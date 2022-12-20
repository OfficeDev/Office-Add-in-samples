namespace AttachmentDemo_ASPNETCore.Models
{
    public class SaveAttachmentsRequestModel
    {
        public string[]? attachmentIds { get; set; }
        public string? messageId { get; set; }
    }
}