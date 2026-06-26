using global::Azure;
using global::Azure.AI.FormRecognizer.DocumentAnalysis;
using Elpis_CRM.Model;
using Microsoft.Extensions.Options;

namespace Elpis_CRM.Services;

/// <summary>
/// Bound configuration (from the "AzureDocumentIntelligence" section) holding the endpoint and API key
/// used to construct the Document Intelligence client.
/// </summary>
public sealed class AzureDocumentIntelligenceSettings
{
    public const string SectionName = "AzureDocumentIntelligence";
    public string Endpoint { get; set; } = string.Empty;
    public string Key { get; set; } = string.Empty;
}

/// <summary>
/// Bound configuration (from the "FileUploadSettings" section) defining the upload size cap and the
/// image content types accepted for scanning. Defaults to 5 MB and common image formats.
/// </summary>
public sealed class FileUploadSettings
{
    public const string SectionName = "FileUploadSettings";
    public long MaxFileSizeBytes { get; set; } = 5 * 1024 * 1024;
    public string[] AllowedContentTypes { get; set; } =
    [
        "image/jpeg", "image/png", "image/tiff", "image/bmp", "image/webp"
    ];
}

/// <summary>
/// Wraps Azure Document Intelligence to turn an uploaded business card image into an <see cref="AccountModel"/>,
/// using the custom "business_card_c1" model and mapping company name, website, phone, and address.
/// </summary>
public sealed class AzureBusinessCardService
{
    private const string ModelId = "business_card_c1";

    private readonly DocumentAnalysisClient _client;
    private readonly FileUploadSettings _uploadSettings;

    /// <summary>
    /// Initializes the service with the Azure analysis client and the upload validation settings.
    /// </summary>
    /// <param name="client">Document Analysis client targeting the configured Azure endpoint.</param>
    /// <param name="uploadSettings">Size and content-type limits applied to incoming files.</param>
    public AzureBusinessCardService(
        DocumentAnalysisClient client,
        IOptions<FileUploadSettings> uploadSettings)
    {
        _client = client;
        _uploadSettings = uploadSettings.Value;
    }

    /// <summary>
    /// Validates the upload, submits it to the Azure "business_card_c1" model, and maps the first detected
    /// document's fields onto a new <see cref="AccountModel"/>.
    /// </summary>
    /// <param name="file">Uploaded image to analyze; checked for emptiness, size, and allowed content type.</param>
    /// <param name="cancellationToken">Token forwarded to the Azure analysis call.</param>
    /// <returns>
    /// An account populated from the card. If Azure detects no document, an account with an empty name is returned;
    /// individual undetected fields are left null.
    /// </returns>
    /// <exception cref="ArgumentException">The file is empty, exceeds the size limit, or has a disallowed content type.</exception>
    /// <exception cref="InvalidOperationException">The Azure Document Intelligence request failed.</exception>
    public async Task<AccountModel> ScanAsync(
        IFormFile file,
        CancellationToken cancellationToken = default)
    {
        if (file.Length == 0)
            throw new ArgumentException("Uploaded file is empty.", nameof(file));

        if (file.Length > _uploadSettings.MaxFileSizeBytes)
            throw new ArgumentException(
                $"File size {file.Length:N0} bytes exceeds the maximum allowed " +
                $"size of {_uploadSettings.MaxFileSizeBytes:N0} bytes.", nameof(file));

        if (!_uploadSettings.AllowedContentTypes.Contains(
                file.ContentType, StringComparer.OrdinalIgnoreCase))
            throw new ArgumentException(
                $"Content type '{file.ContentType}' is not permitted. " +
                $"Allowed: {string.Join(", ", _uploadSettings.AllowedContentTypes)}.", nameof(file));

        AnalyzeResult result;
        try
        {
            await using var stream = file.OpenReadStream();

            var operation = await _client.AnalyzeDocumentAsync(
                WaitUntil.Completed,
                ModelId,
                stream,
                cancellationToken: cancellationToken);

            result = operation.Value;
        }
        catch (RequestFailedException ex)
        {
            throw new InvalidOperationException(
                $"Azure Document Intelligence analysis failed: {ex.Message}", ex);
        }

        if (result.Documents.Count == 0)
            return new AccountModel { Name = string.Empty };

        var doc = result.Documents[0];
        
        // Extract name components
        //var firstName = GetStringField(doc, "FirstName") ?? string.Empty;
        //var lastName = GetStringField(doc, "LastName") ?? string.Empty;
        //var Name= $"{firstName} {lastName}".Trim();

        // ── Debug: Log all field names from Azure response ──
        if (doc.Fields.Count > 0)
        {
            var fieldNames = string.Join(", ", doc.Fields.Keys);
            Console.WriteLine($"[AzureBusinessCardService] Azure response field names: {fieldNames}");
        }
        else
        {
            Console.WriteLine("[AzureBusinessCardService] No" +
                " fields found in Azure response.");
        }

        var account = new AccountModel
        {
            // Name  
            Name = GetFirstListItem(doc,"Company Name"),

            // Company information
            Website = GetFirstListItem(doc, "Website"),
            Phone = GetFirstListItem(doc, "Phone"),

            // Address information
            Address = GetAddressField(doc, "Address"),
            //OwnerEmail = GetFirstListItem(doc,"Email"),

            // Set timestamps
            CreatedAt = DateTime.Now,
            UpdatedAt = DateTime.Now,
            //Active = "Yes"
        };

        return account;
    }


    /// <summary>
    /// Reads a named field as a trimmed string.
    /// </summary>
    /// <param name="doc">Analyzed document returned by Azure.</param>
    /// <param name="fieldName">Key of the field to read.</param>
    /// <returns>The trimmed value, or null if the field is absent or blank.</returns>
    private static string? GetStringField(AnalyzedDocument doc, string fieldName)
    {
        if (!doc.Fields.TryGetValue(fieldName, out var field) || field is null)
            return null;

        var value = field.Value?.AsString();
        return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
    }

    /// <summary>
    /// Reads the first entry of a list-typed field, preferring its phone-number representation and falling back
    /// to its string value; if the field is not a list, reads it directly as a string.
    /// </summary>
    /// <param name="doc">Analyzed document returned by Azure.</param>
    /// <param name="fieldName">Key of the field to read.</param>
    /// <returns>The trimmed value, or null if the field is absent, empty, or unreadable.</returns>
    private static string? GetFirstListItem(AnalyzedDocument doc, string fieldName)
    {
        if (!doc.Fields.TryGetValue(fieldName, out var field) || field is null)
            return null;

        // Try as List first
        try
        {
            var list = field.Value?.AsList();
            if (list is not null && list.Count > 0)
            {
                var item = list[0];
                // Try PhoneNumber first (for phone fields), then fall back to String.
                try
                {
                    var phone = item.Value?.AsPhoneNumber();
                    if (!string.IsNullOrWhiteSpace(phone))
                        return phone.Trim();
                }
                catch (InvalidOperationException) { }

                try
                {
                    var value = item.Value?.AsString();
                    return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
                }
                catch (InvalidOperationException)
                {
                    return null;
                }
            }
        }
        catch (InvalidOperationException) { }

        // Fallback: Try as String
        try
        {
            var value = field.Value?.AsString();
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }
        catch (InvalidOperationException)
        {
            return null;
        }
    }

    /// <summary>
    /// Reads an address field, joining the structured address components (street, city, state, postal code,
    /// country) into a single comma-separated string and falling back to the raw string value when no
    /// structured address is available.
    /// </summary>
    /// <param name="doc">Analyzed document returned by Azure.</param>
    /// <param name="fieldName">Key of the address field to read.</param>
    /// <returns>The formatted address, or null if the field is absent or yields no usable value.</returns>
    private static string? GetAddressField(AnalyzedDocument doc, string fieldName)
    {
        if (!doc.Fields.TryGetValue(fieldName, out var field) || field is null)
            return null;

        // Try as List first
        try
        {
            var list = field.Value?.AsList();
            if (list is not null && list.Count > 0)
            {
                var firstAddr = list[0];
                // Try structured address object first.
                try
                {
                    var structured = firstAddr.Value?.AsAddress();
                    if (structured is not null)
                    {
                        var parts = new[]
                        {
                            structured.StreetAddress,
                            structured.City,
                            structured.State,
                            structured.PostalCode,
                            structured.CountryRegion
                        }.Where(p => !string.IsNullOrWhiteSpace(p));

                        var formatted = string.Join(", ", parts);
                        return string.IsNullOrWhiteSpace(formatted) ? null : formatted;
                    }
                }
                catch (InvalidOperationException) { }

                // Raw string fallback.
                var raw = firstAddr.Value?.AsString();
                return string.IsNullOrWhiteSpace(raw) ? null : raw.Trim();
            }
        }
        catch (InvalidOperationException) { }

        // Fallback: Try as String
        try
        {
            var value = field.Value?.AsString();
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }
        catch (InvalidOperationException)
        {
            return null;
        }
    }
}
