namespace AdactaInternational.AdactaReportsShoppingBag.Model

open System.Xml.Serialization

[<XmlType("Prodotto")>]
type Product = {
    [<XmlAttribute("sigla")>]
    Code: string
    
    [<XmlAttribute("nome")>]
    Name: string
    
    [<XmlAttribute("categoria")>]
    Category: string
    
    [<XmlAttribute("produttore")>]
    Manufacturer: string
    
    [<XmlAttribute("formato")>]
    Format: string
    
    [<XmlAttribute("EAN")>]
    EAN: string
    
    [<XmlAttribute("lotto")>]
    Batch: string
    
    [<XmlAttribute("dataDiProduzione")>]
    ProductionDate: string
    
    [<XmlAttribute("dataDiScadenza")>]
    ExpiryDate: string
    
    [<XmlAttribute("stabilimentoDiProduzione")>]
    ProductionFacility: string
    
    [<XmlElement("ingredienti")>]
    Ingredients: string
    
    [<XmlArray("foto")>]
    [<XmlArrayItem("FotoProdotto")>]
    ProductPhotos: ProductPhoto list
}