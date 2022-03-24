using System;
using System.Collections.Generic;
using System.IO;
using System.ServiceProcess;
using System.Timers;
using Newtonsoft.Json;
using RestSharp;
using EtsyService.Models;
using System.Net;
using System.Threading;
using System.Net.Mail;
using Newtonsoft.Json.Linq;

namespace EtsyService
{
    public partial class Service1 : ServiceBase
    {
        System.Timers.Timer timer = new System.Timers.Timer(); // name space(using System.Timers;)  
        static string getDataStart = "0\":";
        static int pageLimit = 100;
        private static int pollingTime = 1000 * 60 * 60 * 6; //Every 24 Hours
        public static int requestCounter = 0;
        public static string Listings = "";
        public Service1()
        {
            InitializeComponent();
        }
        protected override void OnStart(string[] args)
        {
            WriteToFile("Service is started at " + DateTime.Now);
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = 5000;
            timer.Enabled = true;
        }
        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
        }
        public void OnDebug()
        {
            OnStart(null);
        }

        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            WriteToFile("Starting Application @ " + DateTime.Now);
            timer.Interval = pollingTime;
            try
            {
                Listings = "";
                PollActiveListings();
                PollSoldOutListings();
                PollExpiredListings();
                PollInActiveListings();
                sendMail();
            }
            catch (Exception ex)
            {
                sendMail(false, ex.Message);
            }
        }
        public static void WriteToFile(string Message = "")
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }

        static void preventIdlingOut()
        {
            int totalTime = 0;
            int idleTime = 1000 * 60 * 15;
            while (true)
            {
                totalTime += idleTime;
                if (totalTime >= pollingTime)
                    break;
                Thread.Sleep(idleTime);
                for (int i = 0; i < 10000; i++) ;
            }
        }
        static void PollActiveListings()
        {
            WriteToFile();
            WriteToFile("Active State Polling Started");
            int pageNo = 1;
            while (true)
            {
                var client = new RestClient();
                ShopListings shopListing = new ShopListings();
                client.BaseUrl = new System.Uri("https://openapi.etsy.com/v2/shops/maahira/listings/active?limit=" + pageLimit + "&page=" + pageNo++);
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "OAuth " + OAuthSignatureGenerator.GetAuthorizationHeaderValue(client.BaseUrl, "", "GET"));
                IRestResponse response = client.Execute(request);
                CheckRequestThrottleLimit();
                shopListing = JsonConvert.DeserializeObject<ShopListings>(response.Content);
                if (shopListing.count != null && shopListing.count > 0 && shopListing.results.Count > 0)
                {
                    foreach (Listing listingItem in shopListing.results)
                    {
                        if (IsSkuPresent(listingItem) || listingItem.sku.Count>1)
                        {
                            if (listingItem.sku.Count > 1)
                            {
                                if (listingItem.state == "edit" || listingItem.state == "active")
                                {
                                    if (listingItem.state == "edit")
                                        changeInventoryState(listingItem.listing_id, "active");
                                    updateMultiItemInventory(listingItem.listing_id);
                                }
                            }
                            else
                            {
                                //if (listingItem.sku[0] != "14YC238") continue;
                                //if (!listingItem.description.ToLower().Contains("sku"))
                                //{
                                //    updateDescriptionWithSku(listingItem);
                                //    continue;
                                //}listingItem.sku[0]
                                //continue;
                                var client1 = new RestClient("https://www.silvercityonline.com/stock/src/scripts/getItemData.php?perPage=50&page=1&itemNo=" + listingItem.sku[0] + "&sdt=0000-00-00&edt=0000-00-00");
                                var request1 = new RestRequest(Method.GET);
                                IRestResponse response1 = client1.Execute(request1);
                                CheckRequestThrottleLimit();
                                if (response1.Content.Contains(getDataStart))
                                {
                                    int startInd = response1.Content.IndexOf(getDataStart);
                                    string substr = response1.Content.Substring(startInd + getDataStart.Length);
                                    ItemData stockItem = new ItemData();
                                    stockItem = JsonConvert.DeserializeObject<ItemData>(substr.Substring(0, substr.IndexOf("}") + 1));
                                    if (listingItem.state == "edit" && int.Parse(stockItem.curStock) > 0)
                                    {
                                        changeInventoryState(listingItem.listing_id, "active");
                                        updateInventory(listingItem.listing_id, stockItem.sellPrice, stockItem.curStock, stockItem.itemNo);
                                    }
                                    else if (int.Parse(stockItem.curStock) != listingItem.quantity || double.Parse(listingItem.price) != double.Parse(stockItem.sellPrice)*1.15)
                                    {
                                        if (double.Parse(stockItem.sellPrice) == 0)
                                            stockItem.sellPrice = listingItem.price;
                                        if (int.Parse(stockItem.curStock) > 0)
                                            updateInventory(listingItem.listing_id, stockItem.sellPrice, stockItem.curStock, stockItem.itemNo);
                                        else if (listingItem.state != "edit")
                                            changeInventoryState(listingItem.listing_id, "inactive");
                                    }
                                }
                            }
                        }
                    }
                    WriteToFile("Done for Page " + (pageNo - 1) + " @ " + DateTime.Now);
                }
                else
                {
                    pageNo = 1;
                    break;
                }
            }
            WriteToFile();
            WriteToFile("Active State Polling Done");
        }
        static ItemData getItemDataFromPanel(string sku)
        {
            ItemData stockItem = new ItemData();
            var client1 = new RestClient("https://www.silvercityonline.com/stock/src/scripts/getItemData.php?perPage=50&page=1&itemNo=" + sku + "&sdt=0000-00-00&edt=0000-00-00");
            var request1 = new RestRequest(Method.GET);
            IRestResponse response1 = client1.Execute(request1);
            if (response1.Content.Contains(getDataStart))
            {
                int startInd = response1.Content.IndexOf(getDataStart);
                string substr = response1.Content.Substring(startInd + getDataStart.Length);
                stockItem = JsonConvert.DeserializeObject<ItemData>(substr.Substring(0, substr.IndexOf("}") + 1));
            }
            return stockItem;
        }
        static void updateMultiItemInventory(int listingId)
        {
            var client = new RestClient();
            client.BaseUrl = new Uri("https://openapi.etsy.com/v2/listings/" + listingId + "/inventory?write_missing_inventory=true");
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", "OAuth " + OAuthSignatureGenerator.GetAuthorizationHeaderValue(client.BaseUrl, "", "GET"));
            IRestResponse response = client.Execute(request);
            try
            {
                string sku = "";
                var inventoryVariations = JsonConvert.DeserializeObject<GetInventory>(response.Content);
                dynamic data = JObject.Parse(response.Content);

                foreach (var item in data.results.price_on_property.Children())
                    inventoryVariations.price_on_property = item.Value;
                foreach (var item in data.results.quantity_on_property.Children())
                    inventoryVariations.quantity_on_property = item.Value;
                foreach (var item in data.results.sku_on_property.Children())
                    inventoryVariations.sku_on_property = item.Value;

                UpdateInventory1 updateInventoryList = new UpdateInventory1();
                foreach (Product p in inventoryVariations.results.products)
                {
                    ItemData latestProductData = getItemDataFromPanel(p.sku);

                    if (!String.IsNullOrEmpty(latestProductData.itemNo))
                    {
                        if (double.Parse(latestProductData.sellPrice) > 0)
                            p.offerings[0].price = (Convert.ToDouble(latestProductData.sellPrice) * 1.15).ToString();

                        if (int.Parse(latestProductData.curStock) > 0)
                            p.offerings[0].quantity = int.Parse(latestProductData.curStock);
                        else
                            p.offerings[0].quantity = 0;
                    }

                    sku += p.sku + ",";
                    updateInventoryList.Products.Add(p);
                }

                var client1 = new RestClient("https://openapi.etsy.com/v2/listings/" + listingId + "/inventory");
                var request1 = new RestRequest(Method.PUT);
                request1.AddHeader("Content-Type", "application/x-www-form-urlencoded");
                OAuthProperties properties = OAuthSignatureGenerator.GetMultiProductAuthorizationHeaderValue(client1.BaseUrl, JsonConvert.SerializeObject(updateInventoryList.Products), JsonConvert.SerializeObject(inventoryVariations.price_on_property), JsonConvert.SerializeObject(inventoryVariations.quantity_on_property), JsonConvert.SerializeObject(inventoryVariations.sku_on_property), "PUT");
                request1.AddParameter("oauth_consumer_key", properties.oauth_consumer_key);
                request1.AddParameter("oauth_token", properties.oauth_token);
                request1.AddParameter("oauth_signature_method", properties.oauth_signature_method);
                request1.AddParameter("oauth_timestamp", properties.oauth_timestamp);
                request1.AddParameter("oauth_nonce", properties.oauth_nonce);
                request1.AddParameter("oauth_version", properties.oauth_version);
                request1.AddParameter("oauth_signature", properties.oauth_signature);

                request1.AddParameter("products", JsonConvert.SerializeObject(updateInventoryList.Products));
                request1.AddParameter("price_on_property", JsonConvert.SerializeObject(inventoryVariations.price_on_property));
                request1.AddParameter("quantity_on_property", JsonConvert.SerializeObject(inventoryVariations.quantity_on_property));
                request1.AddParameter("sku_on_property", JsonConvert.SerializeObject(inventoryVariations.sku_on_property));

                IRestResponse response1 = client1.Execute(request1);
                CheckRequestThrottleLimit();
                WriteToFile("Stock Item is Updated with Listing ID: " + listingId + " (" + sku + ")");
                Listings = Listings + sku + ", ";
            }
            catch (Exception e)
            {
                WriteToFile("Error for " + listingId + "-----" + e.StackTrace);
            }
        }

        static bool IsSkuPresent(Listing listing)
        {
            try
            {
                if (listing.price == null)
                    return false;
                if (listing.sku.Count >= 1)
                    return true;
                else if (listing.sku.Count == 0)
                {
                    if (listing.description.ToLower().Contains("sku"))
                    {
                        int i = listing.description.ToLower().LastIndexOf("sku");
                        string subDesc = listing.description.Substring(i);
                        i = subDesc.IndexOf("|");
                        subDesc = subDesc.Substring(i + 1);
                        i = subDesc.IndexOf("|");
                        listing.sku.Add(subDesc.Substring(0, i));
                        return true;
                    }
                }
            }
            catch (Exception e)
            {
                return false;
            }
            return false;
        }
        static void updateDescriptionWithSku(Listing listing)
        {
            listing.description = listing.description + "\n\nSku: |" + listing.sku[0] + "|";
            listing.description = EncodeSpecialChars(listing.description.Replace("&quot;", @""""));
            //listing.description = listing.description.Replace("\r", "%0D%0A").Replace("\n", "%0D%0A");
            //listing.description = listing.description.Replace("%", "%25").Replace(",", "%2C").Replace("*", "%2A").Replace("?", "%3F").Replace("\r", "%0D%0A").Replace("\n", "%0D%0A").Replace(":", "%3A");
            var client = new RestClient(@"https://openapi.etsy.com/v2/listings/" + listing.listing_id + "?description=" + listing.description);
            var request = new RestRequest(Method.PUT);
            request.AddHeader("Authorization", "OAuth " + OAuthSignatureGenerator.GetAuthorizationHeaderValue(client.BaseUrl, ""));
            IRestResponse response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK)
            {
                WriteToFile(listing.url);
            }
            CheckRequestThrottleLimit();
        }
        static string EncodeSpecialChars(string s)
        {
            char[] specialChars = { '%', ',', '*', '?', ':', '–', '-', '&' };
            IDictionary<string, string> otherChars = new Dictionary<string, string>()
                {
                {">","%3E" },
                {"$","%24" },
                {"/","%2F" },
                {"+","%2B" },
                    {"&quot;","\"" },
                    {"!","%21" },
                    {"(","%28"},
                    {")", "%29"},
                    {"\"", "%22"},
                    {"\r","%0D%0A"},
                    {"\n","%0D%0A"},
                };

            foreach (var sc in specialChars)
                s = s.Replace(sc.ToString(), System.Web.HttpUtility.UrlEncode(sc.ToString()).ToUpper());
            foreach (var oc in otherChars)
                s = s.Replace(oc.Key, oc.Value);

            return s;
        }
        static void PollInActiveListings()
        {
            WriteToFile();
            WriteToFile("InActive State Polling Started");
            int pageNo = 1;
            while (true)
            {
                var client = new RestClient();
                ShopListings shopListing = new ShopListings();
                client.BaseUrl = new System.Uri("https://openapi.etsy.com/v2/shops/maahira/listings/inactive?limit=" + pageLimit + "&page=" + pageNo++);
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "OAuth " + OAuthSignatureGenerator.GetAuthorizationHeaderValue(client.BaseUrl, "", "GET"));
                IRestResponse response = client.Execute(request);
                CheckRequestThrottleLimit();
                shopListing = JsonConvert.DeserializeObject<ShopListings>(response.Content);
                if (shopListing.count != null && shopListing.count > 0 && shopListing.results.Count > 0)
                {
                    foreach (Listing listingItem in shopListing.results)
                    {
                        if (IsSkuPresent(listingItem))
                        {
                            //if (!listingItem.description.ToLower().Contains("sku"))
                            //{
                            //    updateDescriptionWithSku(listingItem);
                            //    continue;
                            //}
                            //continue;

                            var client1 = new RestClient("https://www.silvercityonline.com/stock/src/scripts/getItemData.php?perPage=50&page=1&itemNo=" + listingItem.sku[0] + "&sdt=0000-00-00&edt=0000-00-00");
                            var request1 = new RestRequest(Method.GET);
                            IRestResponse response1 = client1.Execute(request1);
                            CheckRequestThrottleLimit();
                            if (response1.Content.Contains(getDataStart))
                            {
                                int startInd = response1.Content.IndexOf(getDataStart);
                                string substr = response1.Content.Substring(startInd + getDataStart.Length);
                                ItemData stockItem = new ItemData();
                                stockItem = JsonConvert.DeserializeObject<ItemData>(substr.Substring(0, substr.IndexOf("}") + 1));
                                if (int.Parse(stockItem.curStock) > 0)
                                {
                                    changeInventoryState(listingItem.listing_id, "active");
                                    updateInventory(listingItem.listing_id, stockItem.sellPrice, stockItem.curStock, stockItem.itemNo);
                                }
                            }
                        }
                    }
                    WriteToFile("Done for Page " + (pageNo - 1) + " @ " + DateTime.Now);
                }
                else
                {
                    pageNo = 1;
                    break;
                }
            }
            WriteToFile();
            WriteToFile("InActive State Polling Done");
        }
        static void PollExpiredListings()
        {
            WriteToFile();
            WriteToFile("Expired State Polling Started");
            int pageNo = 1;
            while (true)
            {
                var client = new RestClient();
                ShopListings shopListing = new ShopListings();
                client.BaseUrl = new System.Uri("https://openapi.etsy.com/v2/shops/maahira/listings/expired?limit=" + pageLimit + "&page=" + pageNo++);
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "OAuth " + OAuthSignatureGenerator.GetAuthorizationHeaderValue(client.BaseUrl, "", "GET"));
                IRestResponse response = client.Execute(request);
                CheckRequestThrottleLimit();
                shopListing = JsonConvert.DeserializeObject<ShopListings>(response.Content);
                if (shopListing.count != null && shopListing.count > 0 && shopListing.results.Count > 0)
                {
                    foreach (Listing listingItem in shopListing.results)
                    {
                        if (IsSkuPresent(listingItem))
                        {
                            var client1 = new RestClient("https://www.silvercityonline.com/stock/src/scripts/getItemData.php?perPage=50&page=1&itemNo=" + listingItem.sku[0] + "&sdt=0000-00-00&edt=0000-00-00");
                            var request1 = new RestRequest(Method.GET);
                            IRestResponse response1 = client1.Execute(request1);
                            CheckRequestThrottleLimit();
                            if (response1.Content.Contains(getDataStart))
                            {
                                int startInd = response1.Content.IndexOf(getDataStart);
                                string substr = response1.Content.Substring(startInd + getDataStart.Length);
                                ItemData stockItem = new ItemData();
                                stockItem = JsonConvert.DeserializeObject<ItemData>(substr.Substring(0, substr.IndexOf("}") + 1));
                                if (int.Parse(stockItem.curStock) > 0)
                                {
                                    changeInventoryState(listingItem.listing_id, "active");
                                    updateInventory(listingItem.listing_id, stockItem.sellPrice, stockItem.curStock, stockItem.itemNo);
                                }
                            }
                        }
                    }
                    WriteToFile("Done for Page " + (pageNo - 1) + " @ " + DateTime.Now);
                }
                else
                {
                    pageNo = 1;
                    break;
                }
            }
            WriteToFile();
            WriteToFile("Expired State Polling Done");
        }
        static void PollSoldOutListings()
        {
            WriteToFile();
            WriteToFile("SoldOut State Polling Started");
            int pageNo = 1;
            while (true)
            {
                var client = new RestClient();
                Transactions transactions = new Transactions();
                client.BaseUrl = new System.Uri("https://openapi.etsy.com/v2/shops/maahira/transactions?includes=Listing&limit=" + pageLimit + "&page=" + pageNo++);
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "OAuth " + OAuthSignatureGenerator.GetAuthorizationHeaderValue(client.BaseUrl, "", "GET"));
                IRestResponse response = client.Execute(request);
                CheckRequestThrottleLimit();
                transactions = JsonConvert.DeserializeObject<Transactions>(response.Content);
                if (transactions.count != null && transactions.count > 0 && transactions.results.Count > 0)
                {
                    foreach (TransactionDetails transaction in transactions.results)
                    {
                        if (IsSkuPresent(transaction.Listing) && !String.IsNullOrEmpty(transaction.Listing.state) && transaction.Listing.state == "sold_out")
                        {
                            if (transaction.Listing.sku[0] != "14YC238") continue;

                            var client1 = new RestClient("https://www.silvercityonline.com/stock/src/scripts/getItemData.php?perPage=50&page=1&itemNo=" + transaction.Listing.sku[0] + "&sdt=0000-00-00&edt=0000-00-00");
                            var request1 = new RestRequest(Method.GET);
                            IRestResponse response1 = client1.Execute(request1);
                            CheckRequestThrottleLimit();
                            if (response1.Content.Contains(getDataStart))
                            {
                                int startInd = response1.Content.IndexOf(getDataStart);
                                string substr = response1.Content.Substring(startInd + getDataStart.Length);
                                ItemData stockItem = new ItemData();
                                stockItem = JsonConvert.DeserializeObject<ItemData>(substr.Substring(0, substr.IndexOf("}") + 1));
                                if (int.Parse(stockItem.curStock) > 0)
                                {
                                    changeInventoryState(transaction.Listing.listing_id, "active");
                                    updateInventory(transaction.Listing.listing_id, stockItem.sellPrice, stockItem.curStock, stockItem.itemNo);
                                }
                            }
                        }
                    }
                    WriteToFile("Done for Page " + (pageNo - 1) + " @ " + DateTime.Now);
                }
                else
                {
                    pageNo = 1;
                    break;
                }
            }
            WriteToFile();
            WriteToFile("Sold Out State Polling Done");
        }
        static void CheckRequestThrottleLimit()
        {
            requestCounter++;
            if (requestCounter % 4 == 0)
                Thread.Sleep(1000);
            requestCounter %= 4;
        }
        static void changeInventoryState(int listingId, string state)
        {
            var renewParam = "";
            if (state == "active")
                renewParam = "&renew=true";
            var client = new RestClient("https://openapi.etsy.com/v2/listings/" + listingId + "?state=" + state + renewParam);
            var request = new RestRequest(Method.PUT);
            request.AddHeader("Authorization", "OAuth " + OAuthSignatureGenerator.GetAuthorizationHeaderValue(client.BaseUrl, ""));
            IRestResponse response = client.Execute(request);
            CheckRequestThrottleLimit();
            WriteToFile("Changing State to " + state + " for listing Id: " + listingId);
            //WriteToFile(response.Content);
        }
        static void updateInventory(int listingId, string sellPrice, string currentQty, string sku)
        {
            var client = new RestClient();
            client.BaseUrl = new Uri("https://openapi.etsy.com/v2/listings/" + listingId + "/inventory?api_key=3ptctueuc44gh9e3sny1oix5&write_missing_inventory=true");
            var request = new RestRequest(Method.GET);
            IRestResponse response = client.Execute(request);
            CheckRequestThrottleLimit();
            try
            {
                var inventoryVariations = JsonConvert.DeserializeObject<GetInventory>(response.Content);
                if (inventoryVariations.results.products.Count == 1)
                {
                    List<UpdateInventory> updateInventoryList = new List<UpdateInventory>();
                    UpdateInventory updateInventory = new UpdateInventory();
                    updateInventory.product_id = inventoryVariations.results.products[0].product_id;
                    updateInventory.offerings.Add(new UpdateOffering());
                    updateInventory.offerings[0].offering_id = inventoryVariations.results.products[0].offerings[0].offering_id;
                    updateInventory.offerings[0].price = (Convert.ToDouble(sellPrice) * 1.15).ToString();
                    updateInventory.offerings[0].quantity = int.Parse(currentQty);
                    updateInventory.sku = sku;
                    updateInventoryList.Add(updateInventory);

                    var client1 = new RestClient("https://openapi.etsy.com/v2/listings/" + listingId + "/inventory");
                    var request1 = new RestRequest(Method.PUT);

                    request1.AddHeader("Authorization", "OAuth " + OAuthSignatureGenerator.GetAuthorizationHeaderValue(client1.BaseUrl, JsonConvert.SerializeObject(updateInventoryList)));
                    request1.AddHeader("Content-Type", "application/x-www-form-urlencoded");
                    request1.AddParameter("products", JsonConvert.SerializeObject(updateInventoryList));

                    IRestResponse response1 = client1.Execute(request1);
                    CheckRequestThrottleLimit();
                    WriteToFile("Stock Item is Updated with Listing ID: " + listingId + " (" + sku + ")");
                    Listings = Listings + sku + ", ";
                    //WriteToFile(response1.Content);
                }
            }
            catch (Exception e)
            {
                WriteToFile("Error for " + listingId + "-----" + e.StackTrace);
            }
        }

        static void sampleUpdateAsync(int listingId = 792035687, string sellPrice = "1350", string currentQty = "1")
        {
            var client = new RestClient();
            client.BaseUrl = new System.Uri("https://openapi.etsy.com/v2/listings/" + listingId + "/inventory?api_key=3ptctueuc44gh9e3sny1oix5&write_missing_inventory=true");
            var request = new RestRequest(Method.GET);
            IRestResponse response = client.Execute(request);
            var inventoryVariations = JsonConvert.DeserializeObject<GetInventory>(response.Content);
            if (inventoryVariations.results.products.Count == 1)
            {
                List<UpdateInventory> updateInventoryList = new List<UpdateInventory>();
                UpdateInventory updateInventory = new UpdateInventory();
                updateInventory.product_id = inventoryVariations.results.products[0].product_id;
                updateInventory.offerings.Add(new UpdateOffering());
                updateInventory.offerings[0].offering_id = inventoryVariations.results.products[0].offerings[0].offering_id;
                updateInventory.offerings[0].price = (Convert.ToDouble(sellPrice) * 1.15).ToString();
                updateInventory.offerings[0].quantity = int.Parse(currentQty);
                updateInventoryList.Add(updateInventory);

                var client1 = new RestClient("https://openapi.etsy.com/v2/listings/" + listingId + "/inventory");
                var request1 = new RestRequest(Method.PUT);

                request1.AddHeader("Authorization", "OAuth " + OAuthSignatureGenerator.GetAuthorizationHeaderValue(new Uri("https://openapi.etsy.com/v2/listings/" + listingId + "/inventory"), JsonConvert.SerializeObject(updateInventoryList)));
                request1.AddHeader("Content-Type", "application/x-www-form-urlencoded");
                request1.AddParameter("products", JsonConvert.SerializeObject(updateInventoryList));

                IRestResponse response1 = client1.Execute(request1);
                WriteToFile(response1.Content);
            }
        }

        static void sendMail(bool success = true, string detail = "")
        {
            try
            {
                Config items = new Config();
                using (StreamReader r = new StreamReader(AppDomain.CurrentDomain.BaseDirectory+"\\config.json"))
                {
                    string json = r.ReadToEnd();
                    items = JsonConvert.DeserializeObject<Config>(json);
                }
                if (!items.sendMail && success)
                    return;

                MailMessage msg = new MailMessage();
                var fromAddress = new MailAddress("noreply@silvercityonline.com", "Etsy Service");
                const string fromPassword = "Silvercity@007";

                msg.From = fromAddress;

                foreach (var email in items.mailRecipients)
                    msg.To.Add(email);

                if (success)
                {
                    msg.Subject = "Etsy Refreshed at " + DateTime.UtcNow.ToString();
                    msg.Body = Listings;
                }

                if (!success)
                {
                    msg.Subject = "Etsy Error at " + DateTime.UtcNow.ToString();
                    msg.Body = Listings + "-------------------" + detail;
                }

                var smtp = new SmtpClient
                {
                    Host = "mail.silvercityonline.com",
                    Port = 26,
                    EnableSsl = false,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = true,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
                };
                WriteToFile("Mail Ready at " + DateTime.Now);
                smtp.Send(msg);
                WriteToFile("Mail Sent at " + DateTime.Now);

            }
            catch(Exception e)
            {
                WriteToFile(e.Message + " " + e.InnerException);
            }
        }
    }

    internal class Config
    {
        public bool sendMail;
        public List<string> mailRecipients;
    }
}
