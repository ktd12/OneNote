using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ComponentModel;
using System.Runtime.Serialization;
using System.Xml.Linq;

using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using Microsoft.Live;
using Newtonsoft.Json;
using System.Globalization;
using Microsoft.Phone.Tasks;

namespace CreateOneNotePage
{

   public class clsOneNote: INotifyPropertyChanged
    {
       /// <summary>
       /// These hints are the default text in the Notebook Name and Section Name text boxes
       /// </summary>
        public const string cSectionNameHint = "Enter Section Name";
        public const string cNotebookNameHint = "Enter Notebook Name";
       /// <summary>
       /// The notebook that the user has retrieved
       /// </summary>
        public clsOneNoteNotebook myNoteBook;

       /// <summary>
       /// The list of sections in the notebook the user has retrieved
       /// </summary>
       public List<clsOneNoteSection> SectionList;

        private string _NotebookName;
       /// <summary>
       /// Name of the notebook location for the new page
       /// A Search for a notebook by name will be case-insensitive. This is
       /// done to avoid conflicts when creating a new notebook.  For example
       /// the system will not allow us to create a notebook called "My friends" if
       /// a notebook already exists called "my friends"
       /// </summary>
       public string NotebookName
        {
            get { return _NotebookName; }
            set 
            { 
                _NotebookName = value;
                if (_NotebookName == String.Empty || _NotebookName == cNotebookNameHint)
                { CreateButtonEnabled = false; }
                else
                {
                    if (_SectionName != String.Empty && _SectionName != cSectionNameHint)
                    { CreateButtonEnabled = true; }
                }

                NotifyPropertyChanged("NotebookName");
            }
        }

        private string _SectionName;
        /// <summary>
        /// Name for the section location for the new page
        /// A Search for a section by name will be case-insensitive. This is done to
        /// avoid conflicts when creating a new section.  For example if the notebook has
        /// a section called "Girls" the system will not allow us to create a new section called "girls".
        /// </summary>
        public string SectionName
        {
            get { return _SectionName; }
            set
            {
                _SectionName = value;
                if (_SectionName == String.Empty || SectionName == cSectionNameHint)
                { CreateButtonEnabled = false; }
                else
                {
                    if (_NotebookName != String.Empty && _NotebookName != cNotebookNameHint)
                    { CreateButtonEnabled = true; }
                }

                NotifyPropertyChanged("SectionName");
            }
        }
        

        private string _InfoText;
     /// <summary>
     /// Status of the page creation request
     /// </summary>
        public string InfoText
        {
            get { return _InfoText; }
            set
            {
                if (_InfoText != value)
                {
                    _InfoText = value;
                    NotifyPropertyChanged("InfoText");
                }
            }
        }
        private string _SignedInMsg;
      /// <summary>
      /// Displayed to the user to let them know which account they have signed into
      /// </summary>
        public string SignedInMsg
        {
            get { return _SignedInMsg; }
            set
            {
                if (_SignedInMsg != value)
                {
                    _SignedInMsg = value;
                    NotifyPropertyChanged("SignedInMsg");
                }
            }
        }

        private bool _CreateButtonEnabled;
      /// <summary>
      /// If true, the user can ask to create a page.  Will be false
      /// if the user has not signed in or has not provided a notebook name
      /// and section name
      /// </summary>
        public bool CreateButtonEnabled
        {
            get { return _CreateButtonEnabled; }
            set {
            
                    _CreateButtonEnabled = value;
                    NotifyPropertyChanged("CreateButtonEnabled");               
            }
        }

        private bool _IsSignedIn;
       /// <summary>
       /// True after the user has been authenticated
       /// </summary>
        public bool IsSignedIn
        {
            get { return _IsSignedIn; }
            set
            {

                _IsSignedIn = value;
                NotifyPropertyChanged("IsSignedIn");
            }
        }

        private bool _IsHyperLinkViewNoteVisible;
       
        public bool IsHyperLinkViewNoteVisible
        {
            get { return _IsHyperLinkViewNoteVisible; }
            set
            {
                if (_IsHyperLinkViewNoteVisible != value)
                {
                    _IsHyperLinkViewNoteVisible = value;
                    NotifyPropertyChanged("IsHyperLinkViewNoteVisible");
                }
            }
        }

       
        public string ClientID;
       

        private string _accessToken;
        private DateTimeOffset _accessTokenExpiration;
        private string _refreshToken; // Refresh token (only applicable when the app uses the wl.offline_access wl.signin scopes)
        private StandardResponse _response;

        private const string NotebooksEndpoint = "https://www.onenote.com/api/v1.0/notebooks";
        // OneNote Service API v1.0 Endpoint
        private const string PagesEndpoint = "https://www.onenote.com/api/v1.0/pages";

        // Collateral used to refresh access token (only applicable when the app uses the wl.offline_access wl.signin scopes)
        private const string MsaTokenRefreshUrl = "https://login.live.com/oauth20_token.srf";
        private const string TokenRefreshContentType = "application/x-www-form-urlencoded";
        private const string TokenRefreshRedirectUri = "https://login.live.com/oauth20_desktop.srf";
        private const string TokenRefreshRequestBody = "client_id={0}&redirect_uri={1}&grant_type=refresh_token&refresh_token={2}";

        #region Send create page requests
        /// <summary>
        /// Create a page in a specific section in the DEFAULT Notebook
        /// </summary>
        /// <param name="simpleHtml"></param>
        /// <param name="sectionName"></param>
        /// <returns></returns>
        public async Task CreateSimpleHtml(string simpleHtml, string mySectionName)
        {
            StringBuilder sbMessage = new StringBuilder();
           
            // Create the request message, which is a text/html single part in this case
            // The Service also supports content type multipart/form-data for more complex scenarios

            string LocationOfNote = PagesEndpoint + "/?sectionName=" + mySectionName;
            var createMessage = new HttpRequestMessage(HttpMethod.Post, LocationOfNote)
            {
                Content = new StringContent(simpleHtml, System.Text.Encoding.UTF8, "text/html")
            };

            bool result = await SendCreatePageRequest(createMessage,sbMessage);
        }

        /// <summary>
        /// Create a page in a specific notebook and specific section
        /// Uses the values of the properties NotebookName and SectionName
        /// which are databound to the UI
        /// </summary>
        public async Task CreateSimpleHtml(string simpleHtml)
        {
            StringBuilder sbMessage = new StringBuilder();
            bool result = false;
            clsOneNoteSection tmpSection;
            clsOneNoteNotebook tmpNotebook;
           
            int IndexInList;
            tmpNotebook = new clsOneNoteNotebook { name = NotebookName };
            tmpSection = new clsOneNoteSection { name = SectionName };
            
            CreateButtonEnabled = false;

            InfoText = "Sending page creation request ...";
            //Do we have the correct notebook?
            if (myNoteBook.name != tmpNotebook.name)
            { //Get the Notebook and the section
                result = await GetNoteBook(tmpNotebook,sbMessage);
                if (result == false)
                {
                    IsHyperLinkViewNoteVisible = false;
                    sbMessage.AppendLine("Page Creation Failed. Problem with notebook.");
                    InfoText = sbMessage.ToString();
                    return;
                }

                myNoteBook = tmpNotebook;

                result = await GetSection(myNoteBook, tmpSection,sbMessage);
                if (result == false)
                {
                    IsHyperLinkViewNoteVisible = false;
                    sbMessage.AppendLine("Page creation Failed. Problem with section.");
                    InfoText = sbMessage.ToString();
                    return;
                }
                SectionList.Add(tmpSection);
            }

            IndexInList = SectionList.IndexOf(tmpSection);
            //Do we have the correct Section?
            if (IndexInList < 0)
            { // we don't have the section we need
                if (myNoteBook.name != NotebookName)
                { // we don't have the notebook we need
                    //First get the notebook  or create it if it doesn't exist
                    result = await GetNoteBook(tmpNotebook,sbMessage);
                    if (result == false)
                    {
                        IsHyperLinkViewNoteVisible = false;
                        sbMessage.AppendLine("Page Creation Failed. Problem with notebook.");
                        InfoText = sbMessage.ToString();
                        return;
                    }

                    myNoteBook = tmpNotebook;

                    result = await GetSection(myNoteBook, tmpSection,sbMessage);
                    if (result == false)
                    {
                        IsHyperLinkViewNoteVisible = false;
                        sbMessage.AppendLine("Page creation Failed. Problem with section.");
                        InfoText = sbMessage.ToString();
                        return;
                    }
                    SectionList.Add(tmpSection);
                }
                else
                {
                    result = await GetSection(myNoteBook, tmpSection,sbMessage);
                    if (result==false)
                    {
                        IsHyperLinkViewNoteVisible = false;
                        sbMessage.AppendLine("Page creation Failed. Problem with section.");
                        InfoText = sbMessage.ToString();
                        return;
                    }
                    SectionList.Add(tmpSection);
                }             
            }
            else
            { //we already have this section on our list. How do we know it's in the correct notebook?
                tmpSection = SectionList[IndexInList];
                if (tmpSection.pagesUrl == null)
                {
                    if (tmpNotebook.sectionsUrl == null)
                    {
                        result = await GetNoteBook(tmpNotebook,sbMessage);
                        if (result == false)
                        {
                            IsHyperLinkViewNoteVisible = false;
                            sbMessage.AppendLine("Page Creation Failed. Problem with notebook.");
                            InfoText = sbMessage.ToString();
                            return;
                        }

                        myNoteBook = tmpNotebook;

                        result = await GetSection(myNoteBook, tmpSection,sbMessage);
                        if (result == false)
                        {
                            IsHyperLinkViewNoteVisible = false;
                            sbMessage.AppendLine("Page creation Failed. Problem with section.");
                            InfoText = sbMessage.ToString();
                            return;
                        }
                        SectionList.Add(tmpSection);
                    }
                    else
                    { //get our section from our notebook
                        result = await GetSection(myNoteBook, tmpSection,sbMessage);
                        if (result == false)
                        {
                            IsHyperLinkViewNoteVisible = false;
                            sbMessage.AppendLine("Page creation Failed. Problem with section.");
                            InfoText = sbMessage.ToString();
                            return;
                        }
                        SectionList.Add(tmpSection);
                    }
                }    
            }

         
            // Create the request message, which is a text/html single part in this case
            // The Service also supports content type multipart/form-data for more complex scenarios
            string LocationOfNote = tmpSection.pagesUrl;// +"/?pageName=" + pageName;
            var createMessage = new HttpRequestMessage(HttpMethod.Post, LocationOfNote)
            {
                Content = new StringContent(simpleHtml, System.Text.Encoding.UTF8, "text/html")
            };

            result = await SendCreatePageRequest(createMessage,sbMessage);
            if (result == false)
            {
                IsHyperLinkViewNoteVisible = false;
                sbMessage.AppendLine("Page creation failed");
                InfoText = sbMessage.ToString();
                return;
            }
        }

        /// <summary>
        /// Send a create page request
        /// </summary>
        /// <param name="createMessage">The HttpRequestMessage which contains the page information</param>
        private async Task<bool> SendCreatePageRequest(HttpRequestMessage createMessage, StringBuilder sbMessage)
        {
            var httpClient = new HttpClient();
            _response = null;

            // Check if Auth token needs to be refreshed
            await RefreshAuthTokenIfNeeded();

            // Add Authorization header
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);
            // Note: API only supports JSON return type.
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            // Get and parse the HTTP response from the service
            
            try
            {
                HttpResponseMessage response = await httpClient.SendAsync(createMessage);
                _response = await ParseResponse(response);
                // Update the UI accordingly
                if (UpdateUIAfterCreateRequest(response) == false)
                {
                    if (response.StatusCode == HttpStatusCode.NotFound)
                    { //our local copy of the notebook and section list are no
                        // longer in sync with the user's account. This could
                        // occur because the user deleted the notebook or section
                        sbMessage.AppendLine("Error: Notebook out of synch.  Please exit the app and try again");
                    }
                    else
                    {
                        sbMessage.AppendLine("Error: " + response.StatusCode);
                    }
                    return false; 
                }
            }
            catch (Exception ex)
            {
                sbMessage.AppendLine(ex.Message);
                return false;
            }
                     
            return true;
        }
        #endregion

        #region Sections
        public async Task<bool> GetSection(clsOneNoteNotebook myNotebook, clsOneNoteSection Section, StringBuilder sbMessage)
        {
            bool result = false;
            SectionListResponse mySectionListResponse;
            clsOneNoteSection temp = new clsOneNoteSection{name = SectionName};
            int IndexInList;
            string strRequest;


            strRequest = myNotebook.sectionsUrl;// +"?$filter=contains(name,'" + tmpSection.name + "')";
            var SectionListMessage = new HttpRequestMessage(HttpMethod.Get, strRequest);
            var httpClient = new HttpClient();
            _response = null;

            // Check if Auth token needs to be refreshed
            await RefreshAuthTokenIfNeeded();

            // Add Authorization header
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);
            // Note: API only supports JSON return type.
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            //// Get and parse the HTTP response from the service
            //this.InfoText = AppResources.PageCreationRequest;
            try
            {
                HttpResponseMessage response = await httpClient.SendAsync(SectionListMessage);
                mySectionListResponse = await ParseSectionListResponse(response);
            }
            catch (Exception ex)
            {
                sbMessage.AppendLine(ex.Message);
                return false;
            }
                      
            if (mySectionListResponse.Sections == null || mySectionListResponse.Sections.Count == 0)
            {
                //we didn't find any sections so we need to create ours
               result  = await CreateSection(myNotebook,Section,sbMessage);
            }
            else
            {
                IndexInList = mySectionListResponse.Sections.IndexOf(temp);
                if (IndexInList >= 0)
                {
                    Section.id = mySectionListResponse.Sections[IndexInList].id;
                    Section.name = mySectionListResponse.Sections[IndexInList].name;
                    Section.pagesUrl = mySectionListResponse.Sections[IndexInList].pagesUrl;
                    result = true;
                }
                else
                {
                    //our section is not in the list, we need to create it
                    result = await CreateSection(myNotebook, Section, sbMessage);
                }
           
            }

            return result;
        }

        private async static Task<SectionListResponse> ParseSectionListResponse(HttpResponseMessage response)
        {
            SectionListResponse mySectionListResponse = new SectionListResponse();
            if (response.StatusCode == HttpStatusCode.OK)
            {
                dynamic responseObject = JsonConvert.DeserializeObject(await response.Content.ReadAsStringAsync());

                Newtonsoft.Json.Linq.JArray temp = responseObject.value;

                mySectionListResponse.Sections = temp.Select(p => new clsOneNoteSection
                {
                    name = (string)p["name"],
                    id = (string)p["id"],
                    //oneNoteClientUrl = (string)p["links.oneNoteClientUrl.href"],
                    //oneNoteWebUrl = (string)p["links.oneNoteWebUrl.href"],
                    pagesUrl = (string)p["pagesUrl"]
                }).ToList();
            }

            return mySectionListResponse;
        }
        private async Task<bool> CreateSection(clsOneNoteNotebook Notebook, clsOneNoteSection newSection,StringBuilder sbMessage)
        {
            bool result = false;
            String Body = "{ name: \"" + newSection.name + "\" }";
            HttpRequestMessage createMessage = new HttpRequestMessage(HttpMethod.Post, Notebook.sectionsUrl)
            {
                Content = new StringContent(Body, System.Text.Encoding.UTF8, "application/json")
            };

            try
            {
                result = await SendCreateSectionRequest(createMessage, newSection,sbMessage);
            }
            catch (Exception ex)
            {
                sbMessage.AppendLine(ex.Message);
                return false;
            }
           
            return result;
        }

        private async Task<bool> SendCreateSectionRequest(HttpRequestMessage createMessage,clsOneNoteSection newSection, StringBuilder sbMessage)
        {
            bool result = false;
            var httpClient = new HttpClient();

            CreateSectionResponse successResponse;
            _response = null;

            // Check if Auth token needs to be refreshed
            await RefreshAuthTokenIfNeeded();

            // Add Authorization header
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);
            // Note: API only supports JSON return type.
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            //// Get and parse the HTTP response from the service
          
            try
            {
                HttpResponseMessage response = await httpClient.SendAsync(createMessage);
                _response = await ParseCreateSectionResponse(response);
            }
            catch (Exception ex)
            {
                sbMessage.AppendLine(ex.Message);
                return false;
            }
           
            if (_response as CreateSectionResponse != null)
            {
                successResponse = (CreateSectionResponse)_response;
            }
            else
            { return false; }

            if (successResponse.StatusCode != HttpStatusCode.Created)
            { result = false; }
            else
            {
                newSection.id = successResponse.id;
                newSection.name = successResponse.name;
                newSection.pagesUrl = successResponse.pagesUrl;
                result = true;
            }

            return result;
        }
        /// <summary>
        /// Parse the OneNote Service API create page response
        /// </summary>
        /// <param name="response">The HttpResponseMessage from the create page request</param>
        private async static Task<StandardResponse> ParseCreateSectionResponse(HttpResponseMessage response)
        {
            StandardResponse standardResponse;
            if (response.StatusCode == HttpStatusCode.Created)
            {
                dynamic responseObject = JsonConvert.DeserializeObject(await response.Content.ReadAsStringAsync());
                standardResponse = new CreateSectionResponse
                {
                    StatusCode = response.StatusCode,
                    id = responseObject.id,
                    name= responseObject.name,
                    //self = responseObject.self,
                    //oneNoteClientUrl =  responseObject.links.oneNoteClientUrl.href,
                    //oneNoteWebUrl = responseObject.links.oneNoteWebUrl.href,
                    pagesUrl = responseObject.pagesUrl
                };
            }
            else
            {
                standardResponse = new StandardErrorResponse
                {
                    StatusCode = response.StatusCode,
                    Message = await response.Content.ReadAsStringAsync()
                };
            }

            // Extract the correlation id.  Apps should log this if they want to collect data to diagnose failures with Microsoft support 
            IEnumerable<string> correlationValues;
            if (response.Headers.TryGetValues("X-CorrelationId", out correlationValues))
            {
                standardResponse.CorrelationId = correlationValues.FirstOrDefault();
            }

            return standardResponse;
        }
        #endregion

        #region Notebooks
        public async Task<bool> GetNoteBook(clsOneNoteNotebook Notebook, StringBuilder sbMessage)
        {
            bool result = false;
            NotebookListResponse myNotebookListResponse;
            clsOneNoteNotebook temp = new clsOneNoteNotebook {name = NotebookName };
            int IndexInList;
            string strRequest;


            strRequest = NotebooksEndpoint;// +"?$filter=contains(name,'" + tmpNotebook.name + "')";
          
            var NoteBookListMessage = new HttpRequestMessage(HttpMethod.Get, strRequest);
            var httpClient = new HttpClient();
            _response = null;

            // Check if Auth token needs to be refreshed
            await RefreshAuthTokenIfNeeded();

            // Add Authorization header
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);
            // Note: API only supports JSON return type.
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            //// Get and parse the HTTP response from the service
       
            try
            {
                HttpResponseMessage response = await httpClient.SendAsync(NoteBookListMessage);

                myNotebookListResponse = await ParseNotebookListResponse(response);
            }
            catch (Exception ex)
            {
                sbMessage.AppendLine(ex.Message);
                return false;
            }
          
            if (myNotebookListResponse.Notebooks == null || myNotebookListResponse.Notebooks.Count == 0)
            {
                //we didn't find any notebooks so we need to create ours
                result = await CreateNoteBook(Notebook,sbMessage);
            }
            else
            {
 
                IndexInList = myNotebookListResponse.Notebooks.IndexOf(temp);
                if (IndexInList >= 0 )
                {
                    Notebook.id = myNotebookListResponse.Notebooks[IndexInList].id;
                    Notebook.name = myNotebookListResponse.Notebooks[IndexInList].name;
                    Notebook.sectionsUrl = myNotebookListResponse.Notebooks[IndexInList].sectionsUrl;
                    result = true;
                }
                else
                {
                    //our notebook is not on the list so we need to create it
                    result = await CreateNoteBook(Notebook, sbMessage);
                }                
            }
 
            return result;
        }
        private async Task<bool> CreateNoteBook(clsOneNoteNotebook newNotebook, StringBuilder sbMessage)
        {
            bool result = false;
            String Body = "{ name: \""  + newNotebook.name  +"\" }";// +;
           
            HttpRequestMessage createMessage = new HttpRequestMessage(HttpMethod.Post, NotebooksEndpoint)
            {
                Content = new StringContent(Body, System.Text.Encoding.UTF8, "application/json")
            };
            try
            {
                result = await SendCreateNotebookRequest(createMessage, newNotebook,sbMessage);
            }
            catch (Exception ex)
            {
                sbMessage.AppendLine(ex.Message);
                return false;
            }
          
           return result;
        }
   
        private async Task<bool> SendCreateNotebookRequest(HttpRequestMessage createMessage,clsOneNoteNotebook newNotebook, StringBuilder sbMessage)
        {
            bool result = false;
            var httpClient = new HttpClient();
   
            CreateNotebookResponse successResponse;
            _response = null;

            // Check if Auth token needs to be refreshed
            await RefreshAuthTokenIfNeeded();

            // Add Authorization header
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);
            // Note: API only supports JSON return type.
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            // Get and parse the HTTP response from the service
                    
            try
            {
                HttpResponseMessage response = await httpClient.SendAsync(createMessage);
                _response = await ParseCreateNotebookResponse(response);
                // Update the UI accordingly
                if (UpdateUIAfterCreateRequest(response) == false)
                {
                    sbMessage.AppendLine("Error: " + response.StatusCode);
                    return false; 
                }
            }
            catch (Exception ex)
            {
                sbMessage.AppendLine(ex.Message);
                return false;
            }
        

            if (_response as CreateNotebookResponse != null)
            {
                successResponse = (CreateNotebookResponse)_response;
                
            }
            else
            { return false; }

            if (successResponse.StatusCode != HttpStatusCode.Created)
            { return false; }
            else
            {
                newNotebook.id = successResponse.id;
                newNotebook.name = successResponse.name;
                newNotebook.sectionsUrl = successResponse.sectionsUrl;
                newNotebook.oneNoteClientUrl = successResponse.oneNoteClientUrl;
                newNotebook.oneNoteWebUrl = successResponse.oneNoteWebUrl;

                result = true;
            }

            return result;
        }

        private async static Task<NotebookListResponse> ParseNotebookListResponse(HttpResponseMessage response)
        {
            NotebookListResponse myNotebookListResponse = new NotebookListResponse();
            if (response.StatusCode == HttpStatusCode.OK)
            {
                dynamic responseObject = JsonConvert.DeserializeObject(await response.Content.ReadAsStringAsync());

                Newtonsoft.Json.Linq.JArray temp = responseObject.value;

                myNotebookListResponse.Notebooks = temp.Select(p => new clsOneNoteNotebook
                {
                    name = (string)p["name"],
                    id = (string)p["id"],
                    oneNoteClientUrl = (string)p["links.oneNoteClientUrl.href"],
                    oneNoteWebUrl = (string)p["links.oneNoteWebUrl.href"],
                    sectionsUrl =(string)p["sectionsUrl"]
                }).ToList();
            }
           
            return myNotebookListResponse;
        }
        /// <summary>
        /// Parse the OneNote Service API create page response
        /// </summary>
        /// <param name="response">The HttpResponseMessage from the create page request</param>
        private async static Task<StandardResponse> ParseCreateNotebookResponse(HttpResponseMessage response)
        {
            StandardResponse standardResponse;
            if (response.StatusCode == HttpStatusCode.Created)
            {
                dynamic responseObject = JsonConvert.DeserializeObject(await response.Content.ReadAsStringAsync());
                standardResponse = new CreateNotebookResponse
                {
                    StatusCode = response.StatusCode,
                    id = responseObject.id,
                    name= responseObject.name,
                    self = responseObject.self,
                    oneNoteClientUrl =  responseObject.links.oneNoteClientUrl.href,
                    oneNoteWebUrl = responseObject.links.oneNoteWebUrl.href,
                    sectionsUrl = responseObject.sectionsUrl
                };
            }
            else
            {
                standardResponse = new StandardErrorResponse
                {
                    StatusCode = response.StatusCode,
                    Message = await response.Content.ReadAsStringAsync()
                };
            }

            // Extract the correlation id.  Apps should log this if they want to collect data to diagnose failures with Microsoft support 
            IEnumerable<string> correlationValues;
            if (response.Headers.TryGetValues("X-CorrelationId", out correlationValues))
            {
                standardResponse.CorrelationId = correlationValues.FirstOrDefault();
            }

            return standardResponse;
        }
        #endregion

        #region Open OneNote Hyperlink

        /// <summary>
        /// Open the created page in the OneNote app
        /// </summary>
        public async Task HyperlinkButton_Click()
        {
            if (_response as CreateSuccessResponse != null)
            {
                CreateSuccessResponse successResponse = (CreateSuccessResponse)_response;
                //WebBrowserTask webTask = new WebBrowserTask();
                //webTask.Uri = new Uri(successResponse.OneNoteWebUrl, UriKind.Absolute);
                //webTask.Show();
                //await Windows.System.Launcher.LaunchUriAsync(new Uri(successResponse.OneNoteWebUrl, UriKind.Absolute));
                //Windows.System.Launcher.LaunchUriAsync(new Uri(successResponse.OneNoteWebUrl));

                bool success = await Windows.System.Launcher.LaunchUriAsync(FormulatePageUri(successResponse.OneNoteClientUrl));
            }
        }

        /// <summary>
        /// Formulate the OneNoteClientUrl so that we can open the OneNote app directly
        /// </summary>
        /// <param name="oneNoteClientUrl">The OneNoteClientUrl received in the API JSON response</param>
        private static Uri FormulatePageUri(string oneNoteClientUrl)
        {
            // Regular expression for identifying GUIDs in the URL returned by the server.
            // We need to wrap such GUIDs in curly braces before sending them to OneNote.
            Regex guidRegex = new Regex(@"=([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})&",
                RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
            if (!String.IsNullOrEmpty(oneNoteClientUrl))
            {
                var matches = guidRegex.Matches(oneNoteClientUrl);
                if (matches.Count == 2)
                {
                    oneNoteClientUrl =
                        oneNoteClientUrl.Replace(matches[0].Groups[1].Value, "{" + matches[0].Groups[1].Value + "}")
                            .Replace(matches[1].Groups[1].Value, "{" + matches[1].Groups[1].Value + "}");
                }
                return new Uri(oneNoteClientUrl);
            }
            return null;
        }

        #endregion

        #region LiveSDK authentication/token refresh

        private LiveConnectClient client;

        /// <summary>
        /// This method is called when the Live session status changes
        /// </summary>
        public async Task OnSessionChanged(object sender, Microsoft.Live.Controls.LiveConnectSessionChangedEventArgs e)
        {                 
            switch (e.Status)
            {
                case LiveConnectSessionStatus.Connected:
                    _accessToken = e.Session.AccessToken;
                    _accessTokenExpiration = e.Session.Expires;
                    _refreshToken = e.Session.RefreshToken;

                    this.IsSignedIn = true;

                    client = new LiveConnectClient(e.Session);
                    
                    LiveOperationResult operationResult = await client.GetAsync("me");
                    try
                    {
                        dynamic meResult = operationResult.Result;
                        if (meResult.first_name != null &&
                            meResult.last_name != null)
                        {
                            this.SignedInMsg =
                                meResult.first_name + " " +
                                meResult.last_name + " " + "is signed in";
                        }
                        else
                        {
                            this.SignedInMsg = "Authentication successful";
                        }
                    }
                    catch (LiveConnectException exception)
                    {
                        this.IsSignedIn = false;
                        this.SignedInMsg = "Not Authenticated";// + " " + exception.Message;
                        this.InfoText = "Error calling API: " +
                            exception.Message;                   
                    }
                
                    break;
                case LiveConnectSessionStatus.NotConnected:
                    this.IsSignedIn = false;
                    this.CreateButtonEnabled = false;
                    this.SignedInMsg = "Authentication failed";
                    break;
                default:
                    this.IsSignedIn = false;
                    this.CreateButtonEnabled = false;
                    this.SignedInMsg = "Not Authenticated";
                    break;
            }
         
        }

        /// <summary>
        /// This method tries to refresh the token if it expires. The authentication token needs to be
        /// refreshed continuosly, so that the user is not prompted to sign in again
        /// </summary>
        /// <returns></returns>
        private async Task AttemptAccessTokenRefresh()
        {
            var createMessage = new HttpRequestMessage(HttpMethod.Post, MsaTokenRefreshUrl)
            {
                Content = new StringContent(
                    String.Format(CultureInfo.InvariantCulture, TokenRefreshRequestBody,
                        this.ClientID,
                        TokenRefreshRedirectUri,
                        _refreshToken),
                    System.Text.Encoding.UTF8,
                    TokenRefreshContentType)
            };

            HttpClient httpClient = new HttpClient();
            HttpResponseMessage response = await httpClient.SendAsync(createMessage);
            await ParseRefreshTokenResponse(response);
        }

        /// <summary>
        /// Handle the RegreshToken response
        /// </summary>
        /// <param name="response">The HttpResponseMessage from the TokenRefresh request</param>
        private async Task ParseRefreshTokenResponse(HttpResponseMessage response)
        {
            if (response.StatusCode == HttpStatusCode.OK)
            {
                dynamic responseObject = JsonConvert.DeserializeObject(await response.Content.ReadAsStringAsync());
                _accessToken = responseObject.access_token;
                _accessTokenExpiration = _accessTokenExpiration.AddSeconds((double)responseObject.expires_in);
                _refreshToken = responseObject.refresh_token;
            }
        }

        #endregion

        #region Helper functions

        /// <summary>
        /// Get date in ISO8601 format with local timezone offset
        /// </summary>
        /// <returns>Date as ISO8601 string</returns>
        private static string GetDate()
        {
            return DateTime.Now.ToString("o");
        }

        /// <summary>
        /// Get an asset file packaged with the application and return it as a managed stream
        /// </summary>
        /// <param name="assetFile">The path name of an asset relative to the application package root</param>
        /// <returns>A managed stream of the asset file data, opened for read</returns>
        //private static Stream GetAssetFileStream(string assetFile)
        //{
        //    StreamResourceInfo resource = Application.GetResourceStream(new Uri(assetFile, UriKind.Relative));
        //    return resource.Stream;
        //}

        /// <summary>
        /// Refreshes the live authentication token if it has expired
        /// </summary>
        private async Task RefreshAuthTokenIfNeeded()
        {
            if (_accessTokenExpiration.CompareTo(DateTimeOffset.UtcNow) <= 0)
            {
                this.InfoText = "Access token needs to be refreshed";
                await AttemptAccessTokenRefresh();
            }
          
            IsHyperLinkViewNoteVisible = false;
            
        }

        /// <summary>
        /// Update the UI after a create page request, depending on if it was
        /// successful or not
        /// </summary>
        /// <param name="response">The HttpResponseMessage from the create page request</param>
        private bool UpdateUIAfterCreateRequest(HttpResponseMessage response)
        {
            if (response.StatusCode == HttpStatusCode.Created)
            {
                InfoText = "Page successfully created.";
                IsHyperLinkViewNoteVisible = true;
                CreateButtonEnabled = true;
                return true;
               
            }
            else
            {
                //Our caller will display the error.  We do it this way
                //because we might also encounter an error via a try/catch block
                //and we want to handle the displaying of errors in one place
                IsHyperLinkViewNoteVisible = false;
                CreateButtonEnabled = true;
                return false;
            }
        }

      
        /// <summary>
        /// Parse the OneNote Service API create page response
        /// </summary>
        /// <param name="response">The HttpResponseMessage from the create page request</param>
        private async static Task<StandardResponse> ParseResponse(HttpResponseMessage response)
        {
            StandardResponse standardResponse;
           
            if (response.StatusCode == HttpStatusCode.Created)
            {
                dynamic responseObject = JsonConvert.DeserializeObject(await response.Content.ReadAsStringAsync());
                standardResponse = new CreateSuccessResponse
                {
                    StatusCode = response.StatusCode,
                    OneNoteClientUrl = responseObject.links.oneNoteClientUrl.href,
                    OneNoteWebUrl = responseObject.links.oneNoteWebUrl.href
                };
                
            }
            else
            {
                standardResponse = new StandardErrorResponse
                {
                    StatusCode = response.StatusCode,
                    Message = await response.Content.ReadAsStringAsync()
                };
            }

            // Extract the correlation id.  Apps should log this if they want to collect data to diagnose failures with Microsoft support 
            IEnumerable<string> correlationValues;
            if (response.Headers.TryGetValues("X-CorrelationId", out correlationValues))
            {
                standardResponse.CorrelationId = correlationValues.FirstOrDefault();
            }

            return standardResponse;
        }

   
        public event PropertyChangedEventHandler PropertyChanged;
        protected void NotifyPropertyChanged(String propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (null != handler)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion

      
        public clsOneNote()
        {
            this.CreateButtonEnabled = false;
            this.InfoText = string.Empty;
            this.SignedInMsg = "Not Signed In";
            this.IsSignedIn = false;
            this.IsHyperLinkViewNoteVisible = false;
            this.SectionName = cSectionNameHint;
            this.NotebookName = cNotebookNameHint;

            myNoteBook = new clsOneNoteNotebook();
            SectionList = new List<clsOneNoteSection>();
        }

       public clsOneNote(string myClientID):this()
        {           
            this.ClientID = myClientID;
        }
    }
}
