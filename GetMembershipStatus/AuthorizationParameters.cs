namespace GetMembershipStatus;

internal class AuthorizationParameters
{
    public static string Username = $"leasing";
    public static string Password = $"api_1396*09*14_ls";
    public static string IpPortNumber = $"192.168.22.75:8080";
    public static string Uri_NationalCode = $"http://{IpPortNumber}/tifcoRestApi/api/v1/getMemberStatusByNationalCode?nationalCode=";
    public static string Uri_PersonCode = $"http://{IpPortNumber}/tifcoRestApi/api/v1/getMemberStatus?personCode=";
}