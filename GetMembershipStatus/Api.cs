using Newtonsoft.Json;
using System.Net.Http.Headers;
using static GetMembershipStatus.Program;

namespace GetMembershipStatus;

public class Api
{
    public static async Task<ApiResult> GetData(string queryParam, InputParamType inputParamType)
    {
        using System.Net.Http.HttpClient client = new();

        var userName = AuthorizationParameters.Username;
        var passwd = AuthorizationParameters.Password;
        var IpPort = AuthorizationParameters.IpPortNumber;
        string url = "";
        if (inputParamType == InputParamType.NationalCode)
        {
            url = AuthorizationParameters.Uri_NationalCode + queryParam;
        }
        if (inputParamType == InputParamType.PersonalCode)
        {
            url = AuthorizationParameters.Uri_PersonCode + queryParam;
        }

        var authToken = System.Text.Encoding.ASCII.GetBytes($"{userName}:{passwd}");

        client.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Basic", Convert.ToBase64String(authToken));

        //********************************************************
        System.Net.Http.HttpResponseMessage response = new();
        try
        {
            response = await client.GetAsync(url);
        }
        catch (HttpRequestException exception)
        {
            string ResponseMessage = exception.Message.ToString();
            if (ResponseMessage == $"A socket operation was attempted to an unreachable host. ({IpPort})")
            {

                string exceptionMessage = @"There was a problem to sending the request to Web Service.
    1- Check your internet connection.
    2- Turn off your VPN.
    3- If the problemt is not solved, Call the network Administrator.";

                return new ApiResult
                { StatusCode = -1, ResponseMessage = exceptionMessage, JsonData = { } };
            }
            else
            {
                string otherExceptionMessage = "The connection with the server is disconnected.";
                return new ApiResult
                { StatusCode = -1, ResponseMessage = otherExceptionMessage, JsonData = { } };
            }
        }
        catch (Exception)
        {
            //ex.Message.ToString().Trim(); Log to file
            string exceptionMessage = @"There was a problem to sending the request to Web Service.";
            return new ApiResult { StatusCode = -1, ResponseMessage = exceptionMessage, JsonData = { } };
        }

        //********************************************************

        var statusCode = (int)response.StatusCode;

        ApiResult finalResultApi = new();

        if (statusCode == 201)
        {
            var resultDictionary = await response.Content.ReadAsStringAsync();

            //**********************************************
            Dictionary<string, string>? responseDictionary =
                JsonConvert.DeserializeObject<Dictionary<string, string>>(resultDictionary);
            //**********************************************

            if (responseDictionary != null)
            {
                bool tryResult = responseDictionary.TryGetValue("areaCode", out string? valueFromRegions);

                if (tryResult)
                {
                    if (valueFromRegions != null && int.TryParse(valueFromRegions, out int intRegionId))
                    {
                        try
                        {
                            List<TbRegion> tbRegions = TbRegionList.tbRegions;
                            var region = from reg in tbRegions
                                         where reg.RegionID == intRegionId
                                         select reg;

                            TbRegion r = region.FirstOrDefault() ??
                                new TbRegion() { CenterID = 0, CenterName = "", RegionID = 0, RegionName = "" };


                            responseDictionary?.Add("centerName", r.CenterName.ToString() ?? "");
                            responseDictionary?.Add("cneterId", r.CenterID.ToString() ?? "");

                        }
                        catch (System.IO.FileNotFoundException)
                        {
                            responseDictionary?.Add("centerName", "");
                            responseDictionary?.Add("cneterId", "");
                        }
                    }
                    else
                    {
                        responseDictionary?.Add("centerName", "");
                        responseDictionary?.Add("cneterId", "");
                    }
                }
                else
                {
                    responseDictionary?.Add("centerName", "");
                    responseDictionary?.Add("cneterId", "");
                }

                finalResultApi.StatusCode = 201;
                finalResultApi.ResponseMessage = "Ok";
                finalResultApi.JsonData = responseDictionary ?? default!;

                return finalResultApi;
            }
            else
            {
                string responseMessage = $"Status code is 201 but responseDictionary is null for {queryParam}";

                finalResultApi.StatusCode = statusCode;
                finalResultApi.ResponseMessage = responseMessage;
                finalResultApi.JsonData = responseDictionary ?? default!;

                return finalResultApi;
            }
        }
        else
        {
            string responseResult = await response.Content.ReadAsStringAsync();

            Dictionary<string, string>? contentDictionary;

            contentDictionary = Newtonsoft.Json.JsonConvert
                .DeserializeObject<Dictionary<string, string>?>(responseResult);

            string responseMessage =
                contentDictionary?["message"].ToString() + $"for {queryParam}" ??
                $"Respons form Api: person not found for {queryParam}";

            finalResultApi.StatusCode = statusCode;
            finalResultApi.ResponseMessage = responseMessage;
            finalResultApi.JsonData = contentDictionary ?? default!;

            return finalResultApi;
        }
    }
}