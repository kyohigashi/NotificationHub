using Swashbuckle.Swagger.Annotations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using TRex.Metadata;
using System.Configuration;
using Microsoft.Azure.NotificationHubs;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

namespace NotificationHubAPIApp.Controllers
{
    public class NotificationController : ApiController
    {
        [HttpPost]
        [SwaggerResponse(HttpStatusCode.OK, Type = typeof(NotificationOutcome))]
        [SwaggerResponse(HttpStatusCode.BadRequest, Type = typeof(ConfigurationErrorsException))]
        [Metadata("Send Message (GCM)")]
        public async System.Threading.Tasks.Task<HttpResponseMessage> SendGCMNotification([Metadata("Connection String")]string connectionString, [Metadata("Hub Name")]string hubName, [Metadata("Message")]string message, [Metadata("Toast Tags")]string tags = null)
        {
		
            try
            {
                NotificationHubClient hub = NotificationHubClient.CreateClientFromConnectionString(connectionString, hubName);
                NotificationOutcome result;

                if (!string.IsNullOrEmpty(tags))
                {
                    result = await hub.SendGcmNativeNotificationAsync(message, tags);
                }
                else
                {
                    result = await hub.SendGcmNativeNotificationAsync(message);
                }

                return Request.CreateResponse<NotificationOutcome>(HttpStatusCode.OK, result);
            }
            catch (ConfigurationErrorsException ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.BareMessage);
            }
        }

        [HttpPost]
        [SwaggerResponse(HttpStatusCode.OK, Type = typeof(NotificationOutcome))]
        [SwaggerResponse(HttpStatusCode.BadRequest, Type = typeof(ConfigurationErrorsException))]
        [Metadata("Send Message (MPNS)")]
        public async System.Threading.Tasks.Task<HttpResponseMessage> SendMPNSNotification([Metadata("Connection String")]string connectionString, [Metadata("Hub Name")]string hubName, [Metadata("Toast XML")]string message, [Metadata("Toast Tags")]string tags = null)
        {
            try
            {
                NotificationHubClient hub = NotificationHubClient.CreateClientFromConnectionString(connectionString, hubName);

                NotificationOutcome result;
                if (!string.IsNullOrEmpty(tags))
                {
                    result = await hub.SendMpnsNativeNotificationAsync(message, tags);
                }
                else
                {
                    result = await hub.SendMpnsNativeNotificationAsync(message);
                }

                return Request.CreateResponse<NotificationOutcome>(HttpStatusCode.OK, result);
            }
            catch (ConfigurationErrorsException ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.BareMessage);
            }
        }

		[Metadata("Send Message By Device list with CSV (MPNS)")]
		public async System.Threading.Tasks.Task<HttpResponseMessage> SendNotifcationByDeviceListWithCSV([Metadata("Connection String")]string connectionString, [Metadata("Hub Name")]string hubName, [Metadata("Toast XML")]string message, [Metadata("Toast XML")]string csv, [Metadata("Toast Tags")]string tags = null)
		{
			try
			{
				IList<string> deviceHandles = null;
				NotificationHubClient hub = NotificationHubClient.CreateClientFromConnectionString(connectionString, hubName);
				NotificationOutcome result;
				Notification notification = new AppleNotification(message);
				result = await hub.SendDirectNotificationAsync(notification, deviceHandles);

				//if (!string.IsNullOrEmpty(tags))
				//{
				//	result = await hub.SendMpnsNativeNotificationAsync(message, tags);
				//}
				//else
				//{
				//	result = await hub.SendMpnsNativeNotificationAsync(message);
				//}

				return Request.CreateResponse<NotificationOutcome>(HttpStatusCode.OK, result);
			}
			catch (ConfigurationErrorsException ex)
			{
				return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.BareMessage);
			}
		}

        [HttpPost]
        [SwaggerResponse(HttpStatusCode.OK, Type = typeof(NotificationOutcome))]
        [SwaggerResponse(HttpStatusCode.BadRequest, Type = typeof(ConfigurationErrorsException))]
        [Metadata("Send Message (Windows Native)")]
        public async System.Threading.Tasks.Task<HttpResponseMessage> SendWindowsNativeNotification([Metadata("Connection String")]string connectionString, [Metadata("Hub Name")]string hubName, [Metadata("Toast XML")]string message, [Metadata("Toast Tags")]string tags = null)
        {
            try
            {
                NotificationHubClient hub = NotificationHubClient.CreateClientFromConnectionString(connectionString, hubName);

                NotificationOutcome result;
                if (!string.IsNullOrEmpty(tags))
                {
                    result = await hub.SendWindowsNativeNotificationAsync(message, tags);
                }
                else
                {
                    result = await hub.SendWindowsNativeNotificationAsync(message);
                }

                return Request.CreateResponse<NotificationOutcome>(HttpStatusCode.OK, result);
            }
            catch (ConfigurationErrorsException ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.BareMessage);
            }
        }

        [HttpPost]
        [SwaggerResponse(HttpStatusCode.OK, Type = typeof(NotificationOutcome))]
        [SwaggerResponse(HttpStatusCode.BadRequest, Type = typeof(ConfigurationErrorsException))]
        [Metadata("Send RAW Message (Windows Native)")]
        public async System.Threading.Tasks.Task<HttpResponseMessage> SendWindowsNativeRawNotification([Metadata("Connection String")]string connectionString, [Metadata("Hub Name")]string hubName, [Metadata("Toast RAW")]string message, [Metadata("Toast Tags")]string tags = null)
        {
            try
            {
                NotificationHubClient hub = NotificationHubClient.CreateClientFromConnectionString(connectionString, hubName);
                
                Notification notification = new WindowsNotification(message);
                notification.Headers.Add("X-WNS-Type", "wns/raw");

                NotificationOutcome result;
                if (!string.IsNullOrEmpty(tags))
                {
                    result = await hub.SendNotificationAsync(notification, tags);
                }
                else
                {
                    result = await hub.SendNotificationAsync(notification);
                }

                return Request.CreateResponse<NotificationOutcome>(HttpStatusCode.OK, result);
            }
            catch (ConfigurationErrorsException ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.BareMessage);
            }
        }

        [HttpPost]
        [SwaggerResponse(HttpStatusCode.OK, Type = typeof(NotificationOutcome))]
        [SwaggerResponse(HttpStatusCode.BadRequest, Type = typeof(ConfigurationErrorsException))]
        [Metadata("Send Message (Baidu Native)")]
        public async System.Threading.Tasks.Task<HttpResponseMessage> SendBaiduNativeNotification([Metadata("Connection String")]string connectionString, [Metadata("Hub Name")]string hubName, [Metadata("Message")]string message, [Metadata("Toast Tags")]string tags = null)
        {
            try
            {
                NotificationHubClient hub = NotificationHubClient.CreateClientFromConnectionString(connectionString, hubName);

                NotificationOutcome result;
                if (!string.IsNullOrEmpty(tags))
                {
                    result = await hub.SendBaiduNativeNotificationAsync(message, tags);
                }
                else
                {
                    result = await hub.SendBaiduNativeNotificationAsync(message);
                }

                return Request.CreateResponse<NotificationOutcome>(HttpStatusCode.OK, result);
            }
            catch (ConfigurationErrorsException ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.BareMessage);
            }
        }

        [HttpPost]
        [SwaggerResponse(HttpStatusCode.OK, Type = typeof(NotificationOutcome))]
        [SwaggerResponse(HttpStatusCode.BadRequest, Type = typeof(ConfigurationErrorsException))]
        [Metadata("Send Message (Apple Native)")]
        public async System.Threading.Tasks.Task<HttpResponseMessage> SendAppleNativeNotification([Metadata("Connection String")]string connectionString, [Metadata("Hub Name")]string hubName, [Metadata("JSON Payload")]string message, [Metadata("Toast Tags")]string tags = null)
        {
            try
            {
                NotificationHubClient hub = NotificationHubClient.CreateClientFromConnectionString(connectionString, hubName);
                
                NotificationOutcome result;
                if (!string.IsNullOrEmpty(tags))
                {
                    result = await hub.SendAppleNativeNotificationAsync(message, tags);
                }
                else
                {
                    result = await hub.SendAppleNativeNotificationAsync(message);
                }

                return Request.CreateResponse<NotificationOutcome>(HttpStatusCode.OK, result);
            }
            catch (ConfigurationErrorsException ex)
            {
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex.BareMessage);
            }
        }

		static string ReadExcelFileUrl(string fileUrl)
		{

			WebClient webClient = new WebClient();
			webClient.UseDefaultCredentials = true;
			Stream stream = webClient.OpenRead(fileUrl);

			// Create a spreadsheet document by supplying the file name.
			SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
				Create(stream, SpreadsheetDocumentType.Workbook);

			// Add a WorkbookPart to the document.
			WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
			workbookpart.Workbook = new Workbook();

			// Add a WorksheetPart to the WorkbookPart.
			WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
			worksheetPart.Worksheet = new Worksheet(new SheetData());

			// Add Sheets to the Workbook.
			Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
				AppendChild<Sheets>(new Sheets());

			// Append a new worksheet and associate it with the workbook.
			Sheet sheet = new Sheet()
			{
				Id = spreadsheetDocument.WorkbookPart.
				GetIdOfPart(worksheetPart),
				SheetId = 1,
				Name = "mySheet"
			};
			sheets.Append(sheet);

			// Close the document.
			spreadsheetDocument.Close();

			Console.WriteLine("The spreadsheet document has been created.\nPress a key.");
			Console.ReadKey();
			return "";

		}
    }
}
