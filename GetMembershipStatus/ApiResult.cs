namespace GetMembershipStatus;

public class ApiResult
{
    public int StatusCode { get; set; }
    public string ResponseMessage { get; set; } = "";
    public Dictionary<string, string> JsonData { get; set; } = new Dictionary<string, string>();
}