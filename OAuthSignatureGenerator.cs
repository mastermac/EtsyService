using EtsyService.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace EtsyService
{
    public class OAuthSignatureGenerator
    {
        private const string OAuthConsumerKey = "oauth_consumer_key";
        private const string OAuthNonce = "oauth_nonce";
        private const string OAuthSignatureMethod = "oauth_signature_method";
        private const string OAuthTimestamp = "oauth_timestamp";
        private const string OAuthToken = "oauth_token";
        private const string OAuthVersion = "oauth_version";
        //private const string productValue = "[product";
        private const string consumerKey = "3ptctueuc44gh9e3sny1oix5";
        private const string consumerSecret = "6mn1apdrtk";
        private const string accessToken = "04643ec822a7b7412ce2dac535a494";
        private const string accessTokenSecret = "c17d48d7c3";


        public enum SignatureMethod
        {
            Plaintext,
            HmacSha1
        }
        private static string GetNonce()
        {
            var rand = new Random();
            var nonce = rand.Next(1000000000);
            return nonce.ToString();
        }

        private static string GetTimeStamp()
        {
            var ts = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, 0);
            return Convert.ToInt64(ts.TotalSeconds).ToString();
        }

        public static string GetAuthorizationHeaderValue(Uri uri, string productConfig, string reqType="PUT")
        {
            var nonce = GetNonce();
            var timeStamp = GetTimeStamp();

            var signature = GenerateSignature(uri, reqType, consumerKey, consumerSecret, nonce, timeStamp, OAuthSignatureGenerator.SignatureMethod.HmacSha1, accessToken, accessTokenSecret, "1.0", productConfig);
            var requestParametersForHeader = new List<string>
            {
                "oauth_consumer_key=\"" + consumerKey + "\"",
                "oauth_token=\"" + accessToken + "\"",
                "oauth_signature_method=\"HMAC-SHA1\"",
                "oauth_timestamp=\"" + timeStamp + "\"",
                "oauth_nonce=\"" + nonce + "\"",
                "oauth_version=\"1.0\"",
                "oauth_signature=\"" + Uri.EscapeDataString(signature) + "\""
            };
            return ConcatList(requestParametersForHeader, ",");
        }
        public static OAuthProperties GetMultiProductAuthorizationHeaderValue(Uri uri, string productConfig, string price, string qty, string sku, string reqType = "PUT")
        {
            var nonce = GetNonce();
            var timeStamp = GetTimeStamp();

            var signature = GenerateSignatureMultiProduct(uri, reqType, consumerKey, consumerSecret, nonce, timeStamp, OAuthSignatureGenerator.SignatureMethod.HmacSha1, accessToken, accessTokenSecret, "1.0", productConfig, price, qty, sku);

            return new OAuthProperties()
            {
                oauth_consumer_key = consumerKey,
                oauth_token = accessToken,
                oauth_signature_method = "HMAC-SHA1",
                oauth_timestamp = timeStamp.ToString(),
                oauth_nonce = nonce,
                oauth_version = "1.0",
                oauth_signature = signature
            };
        }

        private static string ConcatList(IEnumerable<string> source, string separator)
        {
            var sb = new StringBuilder();
            foreach (var s in source)
                if (sb.Length == 0)
                {
                    sb.Append(s);
                }
                else
                {
                    sb.Append(separator);
                    sb.Append(s);
                }
            return sb.ToString();
        }
        public static string GenerateSignatureMultiProduct(Uri uri, string httpMethod, string consumerKey, string consumerSecret, string nonce, string timestamp, SignatureMethod signatureMethod, string token = null, string tokenSecret = null, string version = null, string productValue = "", string price = "", string qty = "", string sku = "")
        {
            if (signatureMethod == SignatureMethod.Plaintext) return HttpUtility.UrlEncode(string.Format("{0}&{1}", consumerSecret, tokenSecret));
            List<KeyValuePair<string, string>> parameters = ConvertQueryStringToListOfKvp(uri.Query);
            AddParameter(parameters, OAuthConsumerKey, consumerKey);
            AddParameter(parameters, OAuthNonce, nonce);
            AddParameter(parameters, OAuthSignatureMethod, "HMAC-SHA1");
            AddParameter(parameters, OAuthTimestamp, timestamp);
            if (!String.IsNullOrEmpty(productValue))
                AddParameter(parameters, "products", productValue);
            AddParameter(parameters, "price_on_property", price);
            AddParameter(parameters, "quantity_on_property", qty);
            AddParameter(parameters, "sku_on_property", sku);
            if (!string.IsNullOrEmpty(token)) AddParameter(parameters, OAuthToken, token);
            if (!string.IsNullOrWhiteSpace(version)) AddParameter(parameters, OAuthVersion, version);
            parameters.Sort((x, y) => x.Key == y.Key ? string.Compare(x.Value, y.Value) : string.Compare(x.Key, y.Key));

            var normalizedUrl = string.Format("{0}://{1}{2}{3}", uri.Scheme, uri.Host, (uri.Scheme == "http" && uri.Port == 80) || (uri.Scheme == "https" && uri.Port == 443) ? null : ":" + uri.Port, uri.AbsolutePath);
            if (!String.IsNullOrEmpty(productValue))
            {
                parameters.RemoveAt(parameters.Count - 3);
                parameters.Insert(7, new KeyValuePair<string, string>("products", SpecialUrlEncode(productValue).Replace("+", "%20")));

                //parameters.Add(new KeyValuePair<string, string>("price_on_property", SpecialUrlEncode(price)));
                //parameters.Add(new KeyValuePair<string, string>("quantity_on_property", SpecialUrlEncode(qty)));
                //parameters.Add(new KeyValuePair<string, string>("sku_on_property", SpecialUrlEncode(sku)));
            }
            var normalizedRequestParameters = string.Join(null, parameters.Select(x => "&" + x.Key + "=" + x.Value)).TrimStart('&');

            var signatureData = string.Format("{0}&{1}&{2}", httpMethod.ToUpper(), UrlEncode(normalizedUrl), UrlEncode(normalizedRequestParameters).Replace("%5C", ""));

            return ComputeHasher(consumerSecret, tokenSecret, signatureData);
        }

        public static string GenerateSignature(Uri uri, string httpMethod, string consumerKey, string consumerSecret, string nonce, string timestamp, SignatureMethod signatureMethod, string token = null, string tokenSecret = null, string version = null, string productValue="")
        {
            if (signatureMethod == SignatureMethod.Plaintext) return HttpUtility.UrlEncode(string.Format("{0}&{1}", consumerSecret, tokenSecret));
            List<KeyValuePair<string,string>> parameters = ConvertQueryStringToListOfKvp(uri.Query);
            AddParameter(parameters, OAuthConsumerKey, consumerKey);
            AddParameter(parameters, OAuthNonce, nonce);
            AddParameter(parameters, OAuthSignatureMethod, "HMAC-SHA1");
            AddParameter(parameters, OAuthTimestamp, timestamp);
            if(!String.IsNullOrEmpty(productValue))
                AddParameter(parameters, "products", productValue);
            if (!string.IsNullOrEmpty(token)) AddParameter(parameters, OAuthToken, token);
            if (!string.IsNullOrWhiteSpace(version)) AddParameter(parameters, OAuthVersion, version);
            parameters.Sort((x, y) => x.Key == y.Key ? string.Compare(x.Value, y.Value) : string.Compare(x.Key, y.Key));

            var normalizedUrl = string.Format("{0}://{1}{2}{3}", uri.Scheme, uri.Host, (uri.Scheme == "http" && uri.Port == 80) || (uri.Scheme == "https" && uri.Port == 443) ? null : ":" + uri.Port, uri.AbsolutePath);
            if (!String.IsNullOrEmpty(productValue))
            {
                parameters.RemoveAt(parameters.Count - 1);
                parameters.Add(new KeyValuePair<string, string>("products", SpecialUrlEncode(productValue)));
            }
            var normalizedRequestParameters = string.Join(null, parameters.Select(x => "&" + x.Key + "=" + x.Value)).TrimStart('&');

            var signatureData = string.Format("{0}&{1}&{2}", httpMethod.ToUpper(), UrlEncode(normalizedUrl), UrlEncode(normalizedRequestParameters).Replace("%5C",""));

            return ComputeHasher(consumerSecret, tokenSecret, signatureData);
        }

        private static List<KeyValuePair<string, string>> ConvertQueryStringToListOfKvp(string queryString)
        {
            return Regex.Matches(queryString, @"(\w+)=.*?(?:&|$)").Cast<Match>().Select(x => x.Value.TrimEnd('&').Split('=')).Select(x => new KeyValuePair<string, string>(x[0], x[1])).ToList();
        }

        private static void AddParameter(ICollection<KeyValuePair<string, string>> parameters, string key, string value)
        {
            parameters.Add(new KeyValuePair<string, string>(key, value));
        }

        private static string ComputeHasher(string consumerSecret, string tokenSecret, string signatureData)
        {
            var key = string.Format("{0}&{1}", UrlEncode(consumerSecret), UrlEncode(tokenSecret));
            var hash = new HMACSHA1 { Key = Encoding.ASCII.GetBytes(key) }.ComputeHash(Encoding.ASCII.GetBytes(signatureData));
            return Convert.ToBase64String(hash);
        }

        private static string UrlEncode(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            const string unreservedChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_.~";
            return value.Select(x => x.ToString()).Aggregate((x, y) => x + (unreservedChars.Contains(y.ToString()) ? y : HttpUtility.UrlEncode(y).ToUpper()));
        }
        private static string SpecialUrlEncode(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            var charArray = value.ToCharArray();
            string newString = "";
            for(int i=0;i<charArray.Length;i++)
            {
                if (!Char.IsLetter(charArray[i]))
                    newString += HttpUtility.UrlEncode(charArray[i].ToString()).ToUpper();
                else
                    newString += charArray[i].ToString();
            }
            return newString;
        }
    }
}
