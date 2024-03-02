using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using WordsReplace.Models;

namespace WordsReplace.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class MainController : ControllerBase
    {
        [HttpGet("Index")]
        public IActionResult Index() => Ok("API working...");

        [HttpPost("ReplaceWord")]
        public IActionResult ReplaceWord([FromBody] ReplaceWordRequest request)
        {
            try
            {
                if (!ModelState.IsValid ||
                    string.IsNullOrEmpty(request.PathDocDefault) ||
                    string.IsNullOrEmpty(request.PathNewDoc) ||
                    string.IsNullOrEmpty(request.OldWord) ||
                    string.IsNullOrEmpty(request.NewWord))
                {
                    return BadRequest("Parâmetros incorretos.");
                }

                System.IO.File.Copy(request.PathDocDefault, request.PathNewDoc, true);

                using (WordprocessingDocument doc = WordprocessingDocument.Open(request.PathNewDoc, true))
                {
                    if (doc.MainDocumentPart == null || doc.MainDocumentPart.Document == null || doc.MainDocumentPart.Document.Body == null)
                        return NoContent();

                    foreach (var text in doc.MainDocumentPart.Document.Body.Descendants<Text>())
                    {
                        if (text.Text.Contains(request.OldWord))
                            text.Text = text.Text.Replace(request.OldWord, request.NewWord);
                    }
                }

                return Ok();
            }
            catch (Exception e)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, $"Erro: {e.Message}");
            }
        }

        [HttpPost("Merge")]
        public IActionResult Merge([FromBody] MergeRequest request)
        {
            try
            {
                using (WordprocessingDocument mergedDoc = WordprocessingDocument.Create(request.MergedFilePath, WordprocessingDocumentType.Document))
                {
                    var mainPart = mergedDoc.AddMainDocumentPart();
                    var body = new Body();
                    var document = new Document(body);
                    mainPart.Document = document;

                    using (WordprocessingDocument doc1 = WordprocessingDocument.Open(request.FilePath1, false))
                    using (WordprocessingDocument doc2 = WordprocessingDocument.Open(request.FilePath2, false))
                    {
                        foreach (var element in doc1.MainDocumentPart.Document.Body.Elements())
                        {
                            var clonedElement = element.CloneNode(true);
                            body.Append(clonedElement);
                        }

                        body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

                        // Clonar o conteúdo do segundo arquivo
                        foreach (var element in doc2.MainDocumentPart.Document.Body.Elements())
                        {
                            var clonedElement = element.CloneNode(true);
                            body.Append(clonedElement);
                        }
                    }

                    return Ok();
                }
            }
            catch (Exception e)
            {
                return StatusCode(500, $"Error: {e.Message}");
            }
        }
    }
}
