using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using ProtectDoc.Models;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace ProtectDoc.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;

    public HomeController(ILogger<HomeController> logger)
    {
        _logger = logger;
    }

    public IActionResult Index()
    {
        return View();
    }
    
    public ActionResult EncryptDocument()
    {
        using FileStream fileStream = new(Path.GetFullPath(@"Data/Input_Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using WordDocument document = new(fileStream, FormatType.Docx);
        document.EncryptDocument("syncfusion");
        MemoryStream stream = new();
        document.Save(stream, FormatType.Docx);
        return File(stream, "application/msword", "Encrypted_Document.docx");
    }

    public ActionResult DecryptDocument()
    {
        using FileStream fileStream = new(Path.GetFullPath(@"Data/Input_Template_Encrypted.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using WordDocument document = new(fileStream, FormatType.Docx, "syncfusion");
        document.RemoveEncryption();
        Syncfusion.EJ2.DocumentEditor.WordDocument sfdtDocument = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(document);
        string sfdt = Newtonsoft.Json.JsonConvert.SerializeObject(sfdtDocument);
        sfdtDocument.Dispose();
        return Content(sfdt, "application/json");
    }

    public ActionResult AllowEditingCommentsOnly()
    {
        using FileStream fileStream = new(Path.GetFullPath(@"Data/Input_Template_Protection.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using WordDocument document = new(fileStream, FormatType.Docx);
        document.Protect(ProtectionType.AllowOnlyComments, "password");
        Syncfusion.EJ2.DocumentEditor.WordDocument sfdtDocument = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(document);
        string sfdt = Newtonsoft.Json.JsonConvert.SerializeObject(sfdtDocument);
        sfdtDocument.Dispose();
        return Content(sfdt, "application/json");
    }

    public ActionResult AllowEditingFormFieldsOnly()
    {
        using FileStream fileStream = new(Path.GetFullPath(@"Data/Input_Template_Protection.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using WordDocument document = new(fileStream, FormatType.Docx);
        document.Protect(ProtectionType.AllowOnlyFormFields, "password");
        Syncfusion.EJ2.DocumentEditor.WordDocument sfdtDocument = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(document);
        string sfdt = Newtonsoft.Json.JsonConvert.SerializeObject(sfdtDocument);
        sfdtDocument.Dispose();
        return Content(sfdt, "application/json");
    }

    public ActionResult AllowRevisionOnly()
    {
        using FileStream fileStream = new(Path.GetFullPath(@"Data/Input_Template_Protection.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using WordDocument document = new(fileStream, FormatType.Docx);
        document.Protect(ProtectionType.AllowOnlyRevisions, "password");
        Syncfusion.EJ2.DocumentEditor.WordDocument sfdtDocument = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(document);
        string sfdt = Newtonsoft.Json.JsonConvert.SerializeObject(sfdtDocument);
        sfdtDocument.Dispose();
        return Content(sfdt, "application/json");
    }
    public ActionResult AllowOnlyReading()
    {
        using FileStream fileStream = new(Path.GetFullPath(@"Data/Input_Template_Protection.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using WordDocument document = new(fileStream, FormatType.Docx);
        document.Protect(ProtectionType.AllowOnlyReading, "password");
        Syncfusion.EJ2.DocumentEditor.WordDocument sfdtDocument = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(document);
        string sfdt = Newtonsoft.Json.JsonConvert.SerializeObject(sfdtDocument);
        sfdtDocument.Dispose();
        return Content(sfdt, "application/json");
    }

    public ActionResult RemoveProtection()
    {
        using FileStream fileStream = new(Path.GetFullPath(@"Data/Input_Template_Protected.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using WordDocument document = new(fileStream, FormatType.Docx);
        document.Protect(ProtectionType.NoProtection, "password");
        Syncfusion.EJ2.DocumentEditor.WordDocument sfdtDocument = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(document);
        string sfdt = Newtonsoft.Json.JsonConvert.SerializeObject(sfdtDocument);
        sfdtDocument.Dispose();
        return Content(sfdt, "application/json");
    }

    public ActionResult AddEditableRange()
    {
        WordDocument document = new();
        WSection? section = document.AddSection() as WSection;
        IWParagraph? paragraph = section.HeadersFooters.Header.AddParagraph();
        paragraph = section.AddParagraph();
        paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
        WTextRange? textRange = paragraph.AppendText("Adventure Works Cycles") as WTextRange;
        paragraph.ApplyStyle(BuiltinStyle.Heading1);
        paragraph = section.AddParagraph();
        EditableRangeStart editableRangeStart = paragraph.AppendEditableRangeStart();
        textRange = paragraph.AppendText("Adventure Works Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company. The company manufactures and sells metal and composite bicycles to North American, European and Asian commercial markets. While its base operation is located in Bothell, Washington with 290 employees, several regional sales teams are located throughout their market base.") as WTextRange;
        paragraph.AppendEditableRangeEnd(editableRangeStart);
        paragraph = section.AddParagraph();
        textRange = paragraph.AppendText("In 2000, Adventure Works Cycles bought a small manufacturing plant, Importadores Neptuno, located in Mexico. Importadores Neptuno manufactures several critical subcomponents for the Adventure Works Cycles product line. These subcomponents are shipped to the Bothell location for final product assembly. In 2001, Importadores Neptuno, became the sole manufacturer and distributor of the touring bicycle product group.") as WTextRange;
        document.Protect(ProtectionType.AllowOnlyReading, "password");
        Syncfusion.EJ2.DocumentEditor.WordDocument sfdtDocument = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(document);
        string sfdt = Newtonsoft.Json.JsonConvert.SerializeObject(sfdtDocument);
        sfdtDocument.Dispose();
        return Content(sfdt, "application/json");
    }

    public ActionResult RemoveEditableRange()
    {
        using FileStream fileStream = new(Path.GetFullPath(@"Data/Input_Template_EditableRange.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using WordDocument document = new(fileStream, FormatType.Docx);
        EditableRange editableRange = document.EditableRanges.FindById("0");
        document.EditableRanges.Remove(editableRange);
        Syncfusion.EJ2.DocumentEditor.WordDocument sfdtDocument = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(document);
        string sfdt = Newtonsoft.Json.JsonConvert.SerializeObject(sfdtDocument);
        sfdtDocument.Dispose();
        return Content(sfdt, "application/json");
    }

    public IActionResult Privacy()
    {
        return View();
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}
