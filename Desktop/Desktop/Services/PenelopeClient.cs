using AdactaInternational.AdactaReportsShoppingBag.Model.Soap;
using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Request;
using AdactaInternational.AdactaReportsShoppingBag.Model.Soap.Response;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Net.Http;
using System.Net.Mime;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace AdactaInternational.AdactaReportsShoppingBag.Desktop.Services;

internal sealed class PenelopeClient(IStorageService storageService) : IPenelopeClient
{
    private static readonly HttpClient HttpClient = new()
    {
        BaseAddress = new Uri("https://services.adactainternational.com/reportservices/ws.asmx"),
        Timeout = TimeSpan.FromSeconds(10)
    };

    [RequiresUnreferencedCode("Uses functionality that may break when trimming.")]
    private static async Task<XmlElement?> SerializeToXmlNodeAsync<TRequest>(TRequest request)
    {
        // Create XmlDocument to hold the serialized object
        var xmlDocument = new XmlDocument();

        // Create the serializer
        var serializer = new XmlSerializer(typeof(TRequest));

        // Serialize the class to XmlElement
        using var memoryStream = new MemoryStream();
        await using var xmlWriter = XmlWriter.Create(memoryStream, new XmlWriterSettings
        {
            Indent = false,
            OmitXmlDeclaration = true,
            Encoding = Encoding.UTF8,
            Async = true
        });
        serializer.Serialize(xmlWriter, request);
        await xmlWriter.FlushAsync();

        // Load the serialized XML into XmlDocument
        memoryStream.Position = 0;
        xmlDocument.Load(memoryStream);

        // Return the root element as XmlElement
        return xmlDocument.DocumentElement;
    }

    [RequiresUnreferencedCode("Uses functionality that may break when trimming.")]
    private static async Task<string> SerializeSoapEnvelopeAsync<TRequest>(TRequest request)
    {
        // Create the SOAP envelope
        var envelope = new SoapEnvelope(new SoapBody(await SerializeToXmlNodeAsync(request)));

        // Create the serializer
        var serializer = new XmlSerializer(typeof(SoapEnvelope));
        var namespaces = new XmlSerializerNamespaces();
        namespaces.Add("soap12", "http://www.w3.org/2003/05/soap-envelope");

        // Serialize the class to string
        using var memoryStream = new MemoryStream();
        await using var xmlWriter = XmlWriter.Create(memoryStream, new XmlWriterSettings
        {
            Indent = true,
            Encoding = Encoding.UTF8,
            Async = true
        });
        serializer.Serialize(xmlWriter, envelope, namespaces);
        await xmlWriter.FlushAsync();

        return Encoding.UTF8.GetString(memoryStream.ToArray());
    }

    [RequiresUnreferencedCode("Uses functionality that may break when trimming.")]
    private static async Task<TResponse?> DeserializeSoapEnvelopeAsync<TResponse>(HttpResponseMessage response)
    {
        // Extract the response content
        var content = await response.Content.ReadAsStringAsync();

        // Create the envelope deserializer
        var envelopeSerializer = new XmlSerializer(typeof(SoapEnvelope));
        var namespaces = new XmlSerializerNamespaces();
        namespaces.Add("soap", "http://www.w3.org/2003/05/soap-envelope");

        // Deserialize the string to envelope class
        using var memoryStream = new MemoryStream(Encoding.UTF8.GetBytes(content));
        var envelope = (SoapEnvelope?)envelopeSerializer.Deserialize(memoryStream);

        // Create the action deserializer
        var action = envelope?.SoapBody.SoapAction;

        if (action is null) return default;

        var actionSerializer = new XmlSerializer(typeof(TResponse));

        // Deserialize the action to class
        using var nodeReader = new XmlNodeReader(action);

        return (TResponse?)actionSerializer.Deserialize(nodeReader);
    }

    [RequiresUnreferencedCode("Uses functionality that may break when trimming.")]
    public async Task<IEnumerable<Product>> GetProductsAsync(string jobCode)
    {
        if (!storageService.DoesContainerExist("Credentials"))
            throw new UnauthorizedAccessException("Credentials container does not exist.");

        var username = storageService.FetchData<string>("Credentials", "Username");
        var password = storageService.FetchData<string>("Credentials", "Password");

        if (username is null || password is null) throw new UnauthorizedAccessException("Missing credentials.");

        var xmlContent = await SerializeSoapEnvelopeAsync(new GetJob(jobCode, username, password));

        // Send the SOAP request
        var content = new StringContent(xmlContent, Encoding.UTF8, MediaTypeNames.Application.Soap);
        var response = await HttpClient.PostAsync("", content);

        // Return the deserialized response
        if (!response.IsSuccessStatusCode) throw new HttpRequestException("Failed to fetch products from the server.");

        var deserializedResponse = await DeserializeSoapEnvelopeAsync<GetJobResponse>(response) ??
                                   throw new SerializationException("Failed to deserialize the server response.");

        // If there is no error, return the products, else throw an exception based on the error code
        if (deserializedResponse.GetJobResult.Error is null)
            return deserializedResponse.GetJobResult.Products;

        return deserializedResponse.GetJobResult.Error.Code switch
        {
            -1 => throw new UnauthorizedAccessException("Invalid credentials."),
            -2 => throw new InvalidOperationException("Job not found."),
            _ => throw new HttpRequestException("An unknown error occurred while fetching products.")
        };
    }
}