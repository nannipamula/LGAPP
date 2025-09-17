using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;

namespace LG_APPWebApi.Controllers
{
    [RoutePrefix("api/templates")]
    public class TemplatesController : ApiController
    {
        // POST api/templates/upload
        [HttpPost]
        [Route("upload")]
        public async Task<IHttpActionResult> Upload([FromBody] UploadTemplateRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.FileBase64))
                return BadRequest("Invalid request: missing file");

            // Decode base64 to bytes
            byte[] fileBytes;
            try
            {
                fileBytes = Convert.FromBase64String(request.FileBase64);
            }
            catch (Exception ex)
            {
                return BadRequest("Invalid base64: " + ex.Message);
            }

            // Validate tokens inside the docx using Open XML SDK
            var validation = ValidateContentControlTags(fileBytes);

            // Compare found tags with payload tokens if provided
            var unknownTags = new List<string>();
            if (request.Tokens != null && request.Tokens.Any())
            {
                var providedTags = request.Tokens.Select(t => t.tag).Where(t => !string.IsNullOrWhiteSpace(t)).ToHashSet(StringComparer.OrdinalIgnoreCase);
                // Tags present in document but not provided
                unknownTags = validation.TagsInDocument.Where(t => !providedTags.Contains(t)).ToList();
            }

            if (unknownTags.Any())
            {
                return Ok(new UploadTemplateResponse
                {
                    Success = false,
                    Message = "Unknown tokens found in document: " + string.Join(", ", unknownTags),
                    Validation = validation
                });
            }

            // Save file with simple versioning
            var templatesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Templates");
            if (!Directory.Exists(templatesPath)) Directory.CreateDirectory(templatesPath);

            var templateId = Guid.NewGuid().ToString("N");
            var version = "1.0";
            var fileName = $"{templateId}_v{version}.docx";
            var filePath = Path.Combine(templatesPath, fileName);

            try
            {
                File.WriteAllBytes(filePath, fileBytes);

                // Save metadata (simple JSON)
                var metadata = new
                {
                    TemplateId = templateId,
                    Version = version,
                    UploadedBy = request.Metadata != null ? request.Metadata.UploadedBy : "unknown",
                    Tenant = request.Metadata != null ? request.Metadata.Tenant : null,
                    UploadedOn = DateTime.UtcNow,
                    WorkflowState = "Draft"
                };
                File.WriteAllText(Path.Combine(templatesPath, templateId + ".json"), JsonConvert.SerializeObject(metadata, Formatting.Indented));

                return Ok(new UploadTemplateResponse
                {
                    Success = true,
                    TemplateId = templateId,
                    Version = version,
                    Workflow = "Draft",
                    Validation = validation
                });
            }
            catch (Exception ex)
            {
                return InternalServerError(ex);
            }
        }

        // Validate content controls tags by reading the docx package
        private ValidationResult ValidateContentControlTags(byte[] docxBytes)
        {
            var result = new ValidationResult();
            using (var ms = new MemoryStream(docxBytes))
            using (var doc = WordprocessingDocument.Open(ms, false))
            {
                var sdtElements = doc.MainDocumentPart.Document.Descendants<SdtElement>();
                foreach (var sdt in sdtElements)
                {
                    var sdtPr = sdt.GetFirstChild<SdtProperties>();
                    if (sdtPr == null) continue;
                    var tag = sdtPr.GetFirstChild<Tag>();
                    var alias = sdtPr.GetFirstChild<Alias>();
                    string tagVal = tag != null ? tag.Val?.Value : null;
                    string aliasVal = alias != null ? alias.Val?.Value : null;
                    if (!string.IsNullOrWhiteSpace(tagVal))
                    {
                        result.TagsInDocument.Add(tagVal);
                    }
                    else if (!string.IsNullOrWhiteSpace(aliasVal))
                    {
                        // Some templates may only have Alias set - include alias as well
                        result.TagsInDocument.Add(aliasVal);
                    }
                }
            }
            result.IsValid = true;
            return result;
        }
    }

    // Models
    public class UploadTemplateRequest
    {
        public string FileBase64 { get; set; }
        public string FileName { get; set; }
        public List<TokenDto> Tokens { get; set; }
        public MetadataDto Metadata { get; set; }
    }

    public class TokenDto
    {
        public string tag { get; set; }
        public string title { get; set; }
        public string type { get; set; }
    }

    public class MetadataDto
    {
        public string UploadedBy { get; set; }
        public string Tenant { get; set; }
    }

    public class UploadTemplateResponse
    {
        public bool Success { get; set; }
        public string TemplateId { get; set; }
        public string Version { get; set; }
        public string Workflow { get; set; }
        public string Message { get; set; }
        public ValidationResult Validation { get; set; }
    }

    public class ValidationResult
    {
        public ValidationResult()
        {
            TagsInDocument = new List<string>();
        }
        public bool IsValid { get; set; }
        public List<string> TagsInDocument { get; set; }
    }
}