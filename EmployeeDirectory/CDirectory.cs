using System.DirectoryServices;
using System.DirectoryServices.Protocols;
using System.Net;



namespace EmployeeDirectory
{
    public class CPerson
    {
        public string UserID { get; set; } = "";
        public string Name { get; set; } = "";
        public string Email { get; set; } = "";
        public string Department { get; set; } = "";
        public string Manager { get; set; } = "";      
        public string Title { get; set; } = "";
        public string BannerID { get; set; } = "";
    }

    public class CPerson2
    {
        public string UserID { get; set; } = "";
        public string FirstName { get; set; } = "";
        public string LastName { get; set; } = "";
        public string Email { get; set; } = "";
        public string Department { get; set; } = "";
        public string Title { get; set; } = "";
        public string Phone { get; set; } = "";
        public string Address { get; set; } = "";
    }

    public class CUser
    {
        public string id { get; set; } = "";
        public string value { get; set; } = "";
        public string label { get; set; } = "";
    }

    public class CDirectory
    {
        static DirectoryEntry? de = null;
        static DirectorySearcher ds = null!;

        public CDirectory()
        {
            de = new DirectoryEntry("LDAP://lbauth.auburn.edu:636")
            {
                Username = "CN=ACESLDP,OU=People,OU=AUMain,DC=auburn,DC=edu",
                Password = "vp4Q^1zG4!KPJumH%lXRe*xm!mUWO7",
                AuthenticationType = AuthenticationTypes.SecureSocketsLayer
            };

            ds = new DirectorySearcher(de)
            {
                SearchScope = System.DirectoryServices.SearchScope.Subtree
            };

        }


        public List<CUser> GetUserSearch(string name)
        {
            ds.Filter = ds?.Filter ?? "(&(objectClass=User) (displayName=" + name + "*))";

            //  ds.PropertiesToLoad.Add("sn");            // surname = last name
            ds.PropertiesToLoad.Add("displayName");
            //      ds.PropertiesToLoad.Add("mail");
            ds.PropertiesToLoad.Add("department");
            //ds.PropertiesToLoad.Add("displayName");
            //     ds.PropertiesToLoad.Add("manager");
            ds.PropertiesToLoad.Add("title");
            ds.PropertiesToLoad.Add("sAMAccountName");

            //SearchResultCollection results = null;
            //SearchResultCollection results = new();
            SearchResultCollection? results = null;

            try
            {
                List<CUser> persons = new List<CUser>();
                results = ds.FindAll();
                foreach (SearchResult result in results)
                {
                    CUser person = new CUser();
                    // person.UserID = userid;

                    string Name = result.Properties["displayName"].Count == 0 ? "" : result.Properties["displayName"][0].ToString();
                    //    string lastname = result.Properties["sn"].Count == 0 ? "" : result.Properties["sn"][0].ToString();
                    //     person.Name = givenname + " " + lastname;
                    //     person.Email = result.Properties["mail"].Count == 0 ? "" : result.Properties["mail"][0].ToString();
                    string Department = result.Properties["department"].Count == 0 ? "" : result.Properties["department"][0].ToString();
                    string Title = result.Properties["title"].Count == 0 ? "" : result.Properties["title"][0].ToString();
                    string UserID = result.Properties["sAMAccountName"][0].ToString();

                    person.label = Name + " (" + UserID + ") " + Title + " / " + Department;
                    person.value = Name + " (" + UserID + ")";

                    persons.Add(person);
                }

                return persons;
            }
            catch (Exception f)
            {
                WriteToEventLog(f.ToString());

                return null!;
            }
            finally
            {
                /*
                if (result != null)
                {
                    //de.Dispose();
                    //ds.Dispose();                    
                }
                */

            }
        }

        public CPerson getUserInfo(string userid)
        {
            ds.Filter = "(&(objectClass=User) (samAccountName=" + userid + "))";

            ds.PropertiesToLoad.Add("sn");            // surname = last name
            ds.PropertiesToLoad.Add("displayName");   // given name or display name ??
            ds.PropertiesToLoad.Add("mail");
            ds.PropertiesToLoad.Add("department");
            //ds.PropertiesToLoad.Add("displayName");
            ds.PropertiesToLoad.Add("manager");
            ds.PropertiesToLoad.Add("title");

            SearchResult? result = null;

            try
            {
                result = ds.FindOne();
                if (result != null)
                {
                    CPerson person = new CPerson();
                    person.UserID = userid;

                    string displayname = result.Properties["displayName"].Count == 0 ? "" : result.Properties["displayName"][0].ToString();
                    string lastname = result.Properties["sn"].Count == 0 ? "" : result.Properties["sn"][0].ToString();
                    //person.Name = givenname + " " + lastname;
                    person.Name = displayname;

                    person.Email = result.Properties["mail"].Count == 0 ? "" : result.Properties["mail"][0].ToString();
                    person.Department = result.Properties["department"].Count == 0 ? "" : result.Properties["department"][0].ToString();
                    person.Title = result.Properties["title"].Count == 0 ? "" : result.Properties["title"][0].ToString();

                    if (result.Properties["manager"].Count > 0)
                    {
                        string[] managers = result.Properties["manager"][0].ToString().Split(',');

                        foreach (string p in managers)
                        {
                            string[] keyvalue = p.Split('=');
                            if (keyvalue[0] == "CN")
                            {
                                person.Manager = keyvalue[1];
                                break;
                            }
                        }
                    }
                    else
                    {
                        person.Manager = "";
                    }

                    /*
                    // BannerID
                    using (AcesRest.Models.AcesEntities5 db = new Models.AcesEntities5())
                    {
                        person.BannerID = (from c in db.Personnels where c.UserID == userid select c.BannerID).FirstOrDefault();
                    }
                    */
                    //    log.WriteToEventLog("[Success] UserID : " + userid, "Directory", "Directory");
                    return person;
                }
                else
                {
                    //    log.WriteToEventLog("[No Result] UserID : " + userid, "Directory", "Directory");
                    return null!;
                }
            }
            catch (Exception f)
            {
                WriteToEventLog(f.ToString());

                return null!;
            }
            finally
            {
                if (result != null)
                {
                    //de.Dispose();
                    //ds.Dispose();                    
                }

            }
        }


        public void getUserInfo2(string userid, out CPerson2 person)
        {
            ds.Filter = "(&(objectClass=User) (samAccountName=" + userid + "))";

            ds.PropertiesToLoad.Add("sn");            // surname = last name
            ds.PropertiesToLoad.Add("givenname");
            ds.PropertiesToLoad.Add("mail");
            ds.PropertiesToLoad.Add("department");
            //ds.PropertiesToLoad.Add("displayName");
            ds.PropertiesToLoad.Add("manager");
            ds.PropertiesToLoad.Add("title");
            ds.PropertiesToLoad.Add("telephoneNumber");
            ds.PropertiesToLoad.Add("physicalDeliveryOfficeName");
            ds.PropertiesToLoad.Add("l");
            ds.PropertiesToLoad.Add("st");
            ds.PropertiesToLoad.Add("postalCode");

            SearchResult? result = null;
            person = new CPerson2();
            try
            {
                result = ds.FindOne();
                if (result != null)
                {
                    person.UserID = userid;
                    person.LastName = result.Properties["sn"].Count == 0 ? "" : result.Properties["sn"][0].ToString();
                    person.FirstName = result.Properties["givenname"].Count == 0 ? "" : result.Properties["givenname"][0].ToString();
                    person.Email = result.Properties["mail"].Count == 0 ? "" : result.Properties["mail"][0].ToString();
                    person.Department = result.Properties["department"].Count == 0 ? "" : result.Properties["department"][0].ToString();
                    person.Title = result.Properties["title"].Count == 0 ? "" : result.Properties["title"][0].ToString();
                    person.Phone = result.Properties["telephoneNumber"].Count == 0 ? "" : result.Properties["telephoneNumber"][0].ToString();
                    person.Address = result.Properties["physicalDeliveryOfficeName"].Count == 0 ? "" : result.Properties["physicalDeliveryOfficeName"][0].ToString();
                    string city = result.Properties["l"].Count == 0 ? "" : result.Properties["l"][0].ToString();
                    string state = result.Properties["st"].Count == 0 ? "" : result.Properties["st"][0].ToString();
                    string zip = result.Properties["postalCode"].Count == 0 ? "" : result.Properties["postalCode"][0].ToString();
                    person.Address = person.Address + " " + city + ", " + state + " " + zip;
                    //return person;
                }
                //else
                //  return null;
            }
            catch (Exception f)
            {
                WriteToEventLog(f.ToString());
                //return null;
            }
            finally
            {
                if (result != null)
                {
                    //de.Dispose();
                    //ds.Dispose();                    
                }

            }
        }

        public bool IsADAuthenticated(string srvr, string usr, string pwd)
        {
            bool authenticated = false;

            try
            {

                LdapDirectoryIdentifier ldap = new LdapDirectoryIdentifier(srvr, 636);
                LdapConnection connection = new LdapConnection(ldap);
                NetworkCredential credential = new NetworkCredential(usr, pwd, "auburn");
                //connection.Credential = credential
                connection.SessionOptions.SecureSocketLayer = true;
                connection.AuthType = AuthType.Negotiate;
                connection.Bind(credential);

                authenticated = true;
            }
            catch (LdapException lexc)
            {
                String error = lexc.ServerErrorMessage;
                WriteToEventLog(error);
            }
            catch (Exception e)
            {
                WriteToEventLog(e.ToString());
            }

            return authenticated;
        }

        static public void WriteToEventLog(string message)
        {
            String filename = "C:\\Log\\Directory\\" + DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") +
                                                          DateTime.Now.Day.ToString("00") + ".txt";

            FileStream fs = new FileStream(filename, FileMode.Append, FileAccess.Write);
            StreamWriter swr = new StreamWriter(fs);
            swr.WriteLine("[" + DateTime.Now + "]" + message);
            swr.Close();
        }
    }
    public void SetFilter(string name)
    {
        #if NET8_0
        ds.Filter = "(&(objectClass=User) (displayName=" + name + "*))";
        #elif NETSTANDARD2_1
        // Code specific to .NET Standard 2.1
        ds.Filter = "(&(objectClass=User) (displayName=" + name + "*))";
        #elif NETCOREAPP3_1
        // Code specific to .NET Core 3.1
        ds.Filter = "(&(objectClass=User) (displayName=" + name + "*))";
        #elif NET481
        // Code specific to .NET Framework 4.8.1
        ds.Filter = "(&(objectClass=User) (displayName=" + name + "*))";
        #endif
    }

}
