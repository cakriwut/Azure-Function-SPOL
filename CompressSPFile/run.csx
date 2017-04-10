#r "System.Runtime"
#r "System.Threading.Tasks"
#r "System.IO.Compression"

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Client;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using System.Linq;
using System.Net;
using System.Runtime;
using System;
using System.IO;
using System.IO.Compression;
using System.Net.Http.Headers;

private static string ClientId = GetEnvironmentVariable("CompressSPFile.ClientId")?? "Unknown";
private static string Cert = GetEnvironmentVariable("CompressSPFile.Cert")?? "Unknown";
private static string CertPassword = GetEnvironmentVariable("CompressSPFile.CertPassword")?? "Unknown";
private static string Authority = GetEnvironmentVariable("CompressSPFile.Authority") ?? "https://unknown.com/badurl";
private static string Resource = GetEnvironmentVariable("CompressSPFile.Resource")?? "Unknown";


public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info("C# HTTP trigger function processed a request.");
    var result = new HttpResponseMessage(HttpStatusCode.OK);
   // parse query parameter
    string siteUrl = req.GetQueryNameValuePairs()
            .FirstOrDefault(q => string.Compare(q.Key, "siteUrl", true) == 0)
            .Value;

    string listTitle = req.GetQueryNameValuePairs()
            .FirstOrDefault(q => string.Compare(q.Key, "listTitle", true) == 0)
            .Value;  

    string itemId = req.GetQueryNameValuePairs()
            .FirstOrDefault(q => string.Compare(q.Key, "itemId", true) == 0)
            .Value;  

    // Get request body
    dynamic data = await req.Content.ReadAsAsync<object>();

    // query string or body data
    siteUrl = siteUrl ?? data?.siteUrl;
    listTitle = listTitle ?? data?.listTitle;
    itemId = itemId ?? data?.itemId;

    var listItemId = int.Parse(itemId ?? "0");    

    using (var ctx = await GetClientContext(siteUrl))
    {
        ctx.Load(ctx.Web.Lists);
        ctx.ExecuteQuery();

        var doclib = ctx.Web.Lists.GetByTitle(listTitle);
        var listItem = doclib.GetItemById(itemId);
        var fileItem = listItem.File;
        var fileBinary = fileItem.OpenBinaryStream();
            
        ctx.Load(fileItem);
        ctx.ExecuteQuery();

        var fName = fileItem.Name;
        using (var memStream = new MemoryStream())
        {
            using (var archive = new ZipArchive(memStream, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry(fName, CompressionLevel.Fastest);
                using (var entryStream = entry.Open())
                {
                    fileBinary.Value.CopyTo(entryStream);
                }
            } 
            log.Info($"Sending zip file ==> { Path.GetFileNameWithoutExtension(fName) }.zip");
            result.Content =  new ByteArrayContent(memStream.ToArray());
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment") 
                { 
                    FileName = $"{Path.GetFileNameWithoutExtension(fName)}.zip" 
                };
        }

    }

    result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
    return result;
}

private async static Task<ClientContext> GetClientContext(string siteUrl)
{
    var authenticationContext = new AuthenticationContext(Authority, false);

    var certPath = Path.Combine(Environment.GetEnvironmentVariable("HOME"), "site\\wwwroot\\CompressSPFile\\", Cert);
    var cert = new X509Certificate2(System.IO.File.ReadAllBytes(certPath),
        CertPassword,
        X509KeyStorageFlags.Exportable |
        X509KeyStorageFlags.MachineKeySet |
        X509KeyStorageFlags.PersistKeySet);

    var authenticationResult = await authenticationContext.AcquireTokenAsync(Resource, new ClientAssertionCertificate(ClientId, cert));
    var token = authenticationResult.AccessToken;

    var ctx = new ClientContext(siteUrl);
    ctx.ExecutingWebRequest += (s, e) =>
    {
        e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + authenticationResult.AccessToken;
    };

    return ctx;
}

public static string GetEnvironmentVariable(string name)
{
    return System.Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
}