using System;
using System.Collections.Generic;
using System.Text;
using TrackWSSample.TrackWebReference;
using System.ServiceModel;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Globalization;


namespace TrackWSSample
{
    class
        TrackWSClient
    {

        static void Main()
        {
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=K:\POS Database\POS Construction_be.mdb";
            string strSQL = "SELECT p_TrackingNumber as UPSTracking, p_Reference5 as SNowRef, UPSTran as ActivityCount FROM TRACKING WHERE (((UPSTran) Is Null) AND ((p_Reference5) <> '') AND ((si_VoidIndicator)='N') AND ((si_ReturnServiceOption)='N')) OR (((UPSTran)<>99) AND ((UPSTran)<>999) AND ((si_VoidIndicator)='N') AND ((p_Reference5) <> '') AND ((si_ReturnServiceOption)='N'))";
            using (OleDbConnection connection = new OleDbConnection(connectionString))             
            {
                OleDbCommand command = new OleDbCommand(strSQL, connection);
                try
                {
                    connection.Open();
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {

                            while (reader.Read())
                            {
                                DateTime now = DateTime.Now;  
                                string strTN = reader.GetString(reader.GetOrdinal("UPSTracking"));
                                Int32 strAC = reader.GetInt32(2); 
                                string strSerNowRef = reader.GetString(reader.GetOrdinal("SNowRef"));

                                try
                                {
                                    TrackService track = new TrackService();
                                    TrackRequest tr = new TrackRequest();
                                    UPSSecurity upss = new UPSSecurity();
                                    UPSSecurityServiceAccessToken upssSvcAccessToken = new UPSSecurityServiceAccessToken();
                                    upssSvcAccessToken.AccessLicenseNumber = "1D5E2960D39CB1B5";
                                    upss.ServiceAccessToken = upssSvcAccessToken;
                                    UPSSecurityUsernameToken upssUsrNameToken = new UPSSecurityUsernameToken();
                                    upssUsrNameToken.Username = "chad jasicki";
                                    upssUsrNameToken.Password = "asdf46#$2";
                                    upss.UsernameToken = upssUsrNameToken;
                                    track.UPSSecurityValue = upss;
                                    RequestType request = new RequestType();
                                    String[] requestOption = { "15" };
                                    request.RequestOption = requestOption;
                                    tr.Request = request;
                                    tr.InquiryNumber = strTN;
                                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12 | System.Net.SecurityProtocolType.Tls | System.Net.SecurityProtocolType.Tls11; //This line will ensure the latest security protocol for consuming the web service call.
                                    TrackResponse trackResponse = track.ProcessTrack(tr);
                                    //Console.WriteLine("The transaction was a " + trackResponse.Response.ResponseStatus.Description);
                                    //Console.WriteLine("Shipment Service " + trackResponse.Shipment[0].Service.Description);

                                    foreach (ShipmentType shipment in trackResponse.Shipment)
                                    {
                                        foreach (PackageType package in shipment.Package)
                                        {
                                            bool del = false;
                                            if (package.TrackingNumber == strTN)
                                            {
                                                int intUPSRecords = package.Activity.Length;
                                                int R = intUPSRecords - strAC;
                                                int N = intUPSRecords;
                                                int i = 0;

                                                foreach (ActivityType Act in package.Activity)
                                                {                                              
                                                    string strdatetime = Act.Date + Act.Time;
                                                    DateTime dt = DateTime.ParseExact(strdatetime, "yyyyMMddHHmmss", CultureInfo.InvariantCulture);
                                                    string StatusCode = Act.Status.Code;
                                                    string strAns = Act.Status.Description + " " + dt + " - " + Act.ActivityLocation.Address.City + ", " + Act.ActivityLocation.Address.StateProvinceCode + " " + Act.ActivityLocation.Address.CountryCode;
                                                    Console.WriteLine(package.TrackingNumber + " " + Act.Status.Description + " " + dt + " - " + Act.ActivityLocation.Address.City + ", " + Act.ActivityLocation.Address.StateProvinceCode + " " + Act.ActivityLocation.Address.CountryCode);
                                                   
                                                    if (R > i)
                                                    {
                                                        N = intUPSRecords - i;
                                                        Console.WriteLine("....Writing Record to UPS Track Table....");
                                                        string strAnsB = strAns.Replace("'", "`");
                                                        strSQL = "INSERT INTO UPSTrack ([UPSTrackNumber], [Status], [RequestNum],[Date],[UPSStatusCode],[Order]) VALUES ('" + strTN + "', '" + strAnsB + "', '" + "RE" + strSerNowRef + "', '" + now + "','" + StatusCode + "','" + N + "')";
                                                        command = new OleDbCommand(strSQL, connection);
                                                        command.ExecuteReader();
                                                        i++;
                                                        strSQL = "UPDATE TRACKING SET UPSTran = " + package.Activity.Length + " WHERE p_TrackingNumber ='" + strTN + "'";
                                                        command = new OleDbCommand(strSQL, connection);
                                                        command.ExecuteReader();
                                                    }
                                                    
                                                    if (Act.Status.Description == "Delivered")
                                                    {
                                                        del = true;
                                                    }                                                   
                                                }
                                            }

                                            if (del == true)
                                            {
                                                // Update rows with 99 when package was delivered
                                                strSQL = "UPDATE TRACKING SET UPSTran = 99 WHERE p_TrackingNumber ='" + package.TrackingNumber + "'";
                                                command = new OleDbCommand(strSQL, connection);
                                                command.ExecuteReader();
                                                del = false;
                                            }
                                        }
                                    }
                                }
                                catch (System.Web.Services.Protocols.SoapException ex)
                                {
                                    Console.WriteLine("");
                                    Console.WriteLine("---------Track Web Service returns error----------------");
                                    Console.WriteLine("---------\"Hard\" is user error \"Transient\" is system error----------------");
                                    Console.WriteLine("SoapException Message= " + ex.Message);
                                    Console.WriteLine("");
                                    Console.WriteLine("SoapException Category:Code:Message= " + ex.Detail.LastChild.InnerText);
                                    Console.WriteLine("");
                                    Console.WriteLine("SoapException XML String for all= " + ex.Detail.LastChild.OuterXml);
                                    Console.WriteLine("");
                                    Console.WriteLine("SoapException StackTrace= " + ex.StackTrace);
                                    Console.WriteLine("-------------------------");
                                    Console.WriteLine("");
                                    Console.WriteLine("Press any Key to Continue");
                                    Console.ReadKey();
                                }
                                catch (System.ServiceModel.CommunicationException ex)
                                {
                                    Console.WriteLine("");
                                    Console.WriteLine("--------------------");
                                    Console.WriteLine("CommunicationException= " + ex.Message);
                                    Console.WriteLine("CommunicationException-StackTrace= " + ex.StackTrace);
                                    Console.WriteLine("-------------------------");
                                    Console.WriteLine("");
                                    Console.WriteLine("Press any Key to Continue");
                                    Console.ReadKey();

                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("");
                                    Console.WriteLine("-------------------------");
                                    Console.WriteLine(" General Exception= " + ex.Message);
                                    Console.WriteLine(" General Exception-StackTrace= " + ex.StackTrace);
                                    Console.WriteLine("-------------------------");
                                    Console.WriteLine(strSerNowRef);
                                    Console.WriteLine("Press any Key to Continue");
                                    Console.ReadKey();
                                }
                                finally
                                {
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("No Records Found");
                            Console.WriteLine("Press any Key to Continue");
                            Console.ReadKey();
                        }
                    }
                    
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Console.WriteLine(" General Exception-StackTrace= " + ex.StackTrace);
                    Console.WriteLine("Press any Key to Continue");
                    Console.ReadKey();
                }
                // The connection is automatically closed becasuse of using block.  
                //Console.ReadKey();
            }
        } 
    }
}

