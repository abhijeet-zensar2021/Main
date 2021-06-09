using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.ServiceModel.Description;
using System.Threading.Tasks;

namespace DataMigrationAzureFunctionApp
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            //int ss = number;

            log.Info("C# HTTP trigger function processed a request.");

            int entityNumber = Convert.ToInt32(req.GetQueryNameValuePairs()
               .FirstOrDefault(q => string.Compare(q.Key, "entityNumber", true) == 0)
               .Value);

            string clientId = "d7e78652-cc5b-42d3-86dd-e21074514da0";
            string clientSecret = "LJbEf3cuG4.6M7_i~swm_qrgP5Kyb.cr47";
            string sourceURL = "https://zrcrmdev1.crm8.dynamics.com/";
            string targetURL = "https://zrrcrmuat1.crm8.dynamics.com/";

            //Connect Source environment
            IOrganizationService serviceSource = GetOrganizationServiceClientSecret(clientId, clientSecret, sourceURL);

            if (serviceSource == null)
            {
                return req.CreateResponse(HttpStatusCode.BadRequest, "Failed to Established Connection on Source environment!!!");
            }

            //Connect Target environment
            IOrganizationService serviceTarget = GetOrganizationServiceClientSecret(clientId, clientSecret, targetURL);
            if (serviceTarget == null)
            {
                return req.CreateResponse(HttpStatusCode.OK, "Failed to Established Connection on Target environment!!!");
            }

            //------------Product Entity----------------------------------
            if (entityNumber == 1)
            {
                EntityCollection unitGroupColl = RetrieveUnitGroupRecords(serviceSource, log);
                //Fetch Product Record
                EntityCollection productColl = RetrieveProductRecords(serviceSource, log);

                //Fetch Unit Records
                EntityCollection unitColl = RetrieveUnitRecords(serviceSource, log);

                //Fetch Price List Records
                EntityCollection priceListColl = RetrievePriceListRecords(serviceSource, log);

                if (productColl == null || productColl.Entities.Count == 0)
                {
                    return req.CreateResponse(HttpStatusCode.OK, "No Product records found to transfer");
                }

                //Create Unit  Records
                //bool successFlagUnit = UnitRecords(unitColl, serviceTarget, log);

                //Create Unit Group Records
                bool successFlagUnitGroup = UnitGroupRecords(unitGroupColl, serviceTarget, log);

                //Create Price list Records
                bool successFlagPriceList = UnitPriceListRecords(priceListColl, serviceTarget, log);

                //Create Product Record
                bool successFlag = CreateProducts(productColl, serviceTarget, log);
            }
            //------------Product Entity----------------------------------*/

            //------------Kb Article----------------------------------
            if (entityNumber == 2)
            {
                EntityCollection kbArticlelateColl = RetrieveKbArticleRecords(serviceSource, log);

                bool successFlagkbArticle = CreatekbArticle(kbArticlelateColl, serviceTarget, log);
            }

            //------------Kb Article----------------------------------

            //------------Type----------------------------------
            if (entityNumber == 3)
            {
                EntityCollection typeColl = RetrieveTypeRecords(serviceSource, log);

                bool successFlagType = CreateType(typeColl, serviceTarget, log);
            }
            //------------Type----------------------------------

            //------------Config Setting----------------------------------
            if (entityNumber == 4)
            {
                EntityCollection configSettingColl = RetrieveConfigSettingRecords(serviceSource, log);

                bool successFlagConfigSetting = CreateConfigSetting(configSettingColl, serviceTarget, log);
            }


            return entityNumber == null
                            ? req.CreateResponse(HttpStatusCode.BadRequest, "Done")
                            : req.CreateResponse(HttpStatusCode.OK, "Successfully Executed..");
        }

        private static bool CreateType(EntityCollection typeColl, IOrganizationService serviceTarget, TraceWriter log)
        {
            EntityCollection typeCollTarget = RetrieveTypeRecords(serviceTarget, log);

            log.Info("Creating Type records.");
            Entity entity = new Entity("hsl_type");
            string entityName = "hsl_type";
            foreach (Entity entObj in typeColl.Entities)
            {
                //check record is present on target envt
                bool flag = checkRecord(typeCollTarget, new Guid(entObj.Attributes["hsl_typeid"].ToString()), entityName); ;

                entity.Id = new Guid(entObj.Attributes["hsl_typeid"].ToString());
                if (entObj.Attributes.Contains("hsl_name"))
                {
                    entity.Attributes["hsl_name"] = entObj.Attributes["hsl_name"].ToString();
                }
                if (entObj.Attributes.Contains("hsl_sortorder"))
                {
                    entity.Attributes["hsl_sortorder"] = Convert.ToInt32(entObj.Attributes["hsl_sortorder"].ToString());
                }
                if (entObj.Attributes.Contains("hsl_purpose"))
                {
                    OptionSetValue optionSet = (OptionSetValue)entObj.Attributes["hsl_purpose"];
                    entity.Attributes["hsl_purpose"] = optionSet;
                }
                if (flag == false)
                {
                    
                    serviceTarget.Create(entity);
                }
                else if (flag==true)
                {
                    serviceTarget.Update(entity);
                }

            }

            return true;
        }

        private static bool checkRecord(EntityCollection typeCollTarget, Guid guidTarget,string entityName)
        
        {
            foreach (Entity entObj in typeCollTarget.Entities)
            {
                Guid source = Guid.Empty;
                if (entityName == "hsl_type")
                {
                     source = new Guid(entObj.Attributes["hsl_typeid"].ToString());
                }
                else if(entityName == "hsl_configsetting")
                {
                    source = new Guid(entObj.Attributes["hsl_configsettingid"].ToString());
                }
                if (source == guidTarget)
                {
                    return true;
                }
            }
            return false;
        }

        private static EntityCollection RetrieveTypeRecords(IOrganizationService serviceSource, TraceWriter log)
        {
            log.Info("Fetching Type Records....");
            EntityCollection collection = null;

            string fetchXML = @"<fetch top='100' distinct='true' >
                              <entity name='hsl_type' >
                                <attribute name='hsl_name' />
                                <attribute name='hsl_sortorder' />
                                <attribute name='hsl_purpose' />
                                <attribute name='hsl_typeid' />
                                
                              </entity>
                            </fetch>";
            collection = serviceSource.RetrieveMultiple(new FetchExpression(fetchXML));

            return collection;
        }

        private static bool CreateConfigSetting(EntityCollection configSettingColl, IOrganizationService serviceTarget, TraceWriter log)
        {
            EntityCollection typeCollTarget = RetrieveConfigSettingRecords(serviceTarget, log);
            log.Info("Creating Config Setting records.");
            Entity entity = new Entity("hsl_configsetting");
            string entityName = "hsl_configsetting";
            foreach (Entity entObj in configSettingColl.Entities)
            {
                //check record is present on target envt
                bool flag = checkRecord(typeCollTarget, new Guid(entObj.Attributes["hsl_configsettingid"].ToString()), entityName);

                entity.Id = new Guid(entObj.Attributes["hsl_configsettingid"].ToString());
                if (entObj.Attributes.Contains("hsl_name"))
                {
                    entity.Attributes["hsl_name"] = entObj.Attributes["hsl_name"].ToString();
                }
                if (entObj.Attributes.Contains("hsl_securevalue"))
                {
                    entity.Attributes["hsl_securevalue"] = entObj.Attributes["hsl_securevalue"].ToString();
                }
                if (flag == false)
                {
                    serviceTarget.Create(entity);
                }
                else if(flag == true)
                {
                    serviceTarget.Update(entity);
                }
            }

            return true;
        }

        private static EntityCollection RetrieveConfigSettingRecords(IOrganizationService serviceSource, TraceWriter log)
        {
            log.Info("Fetching Config Setting Records....");
            EntityCollection collection = null;

            string fetchXML = @"<fetch top='100' distinct='true' >
                              <entity name='hsl_configsetting' >
                                <attribute name='hsl_configsettingid' />
                                <attribute name='hsl_securevalue' />
                                <attribute name='hsl_name' />
                               
                              </entity>
                            </fetch>";
            collection = serviceSource.RetrieveMultiple(new FetchExpression(fetchXML));

            return collection;
        }

        private static bool CreatekbArticle(EntityCollection kbArticlelateColl, IOrganizationService serviceTarget, TraceWriter log)
        {
            log.Info("Creating Article records.");
            Entity entity = new Entity("kbarticle");
            foreach (Entity entObj in kbArticlelateColl.Entities)
            {
                entity.Id = new Guid(entObj.Attributes["kbarticleid"].ToString());

                if (entObj.Attributes.Contains("title"))
                {
                    entity.Attributes["title"] = entObj.Attributes["title"].ToString();
                }
                if (entObj.Attributes.Contains("kbarticletemplateid"))
                {
                    EntityReference entref = (EntityReference)entObj.Attributes["kbarticletemplateid"];
                    var LookupId = entref.Id;
                    entity.Attributes["kbarticletemplateid"] = new EntityReference(entref.LogicalName, LookupId);
                }
                //if (entObj.Attributes.Contains("subjectid"))
                //{
                //    EntityReference entref = (EntityReference)entObj.Attributes["subjectid"];
                //    var LookupId = entref.Id;
                //    entity.Attributes["subjectid"] = new EntityReference(entref.LogicalName, LookupId);
                //}

                serviceTarget.Create(entity);
            }

            return true;
        }

        private static EntityCollection RetrieveKbArticleRecords(IOrganizationService serviceSource, TraceWriter log)
        {
            log.Info("Fetching KB Article Records....");
            EntityCollection collection = null;
            string fetchXML = @"<fetch top='50' >
                                  <entity name='kbarticle' >
                                    <attribute name='createdonbehalfbyyominame' />
                                    <attribute name='modifiedonbehalfby' />
                                    <attribute name='transactioncurrencyidname' />
                                    <attribute name='statecode' />
                                    <attribute name='languagecode' />
                                    <attribute name='createdon' />
                                    <attribute name='statecodename' />
                                    <attribute name='kbarticletemplateidtitle' />
                                    <attribute name='createdonbehalfby' />
                                    <attribute name='transactioncurrencyid' />
                                    <attribute name='entityimage_timestamp' />
                                    <attribute name='entityimage' />
                                    <attribute name='entityimageid' />
                                    <attribute name='importsequencenumber' />
                                    <attribute name='organizationid' />
                                    <attribute name='modifiedbyyominame' />
                                    <attribute name='kbarticletemplateid' />
                                    <attribute name='subjectid' />
                                    <attribute name='comments' />
                                    <attribute name='statuscode' />
                                    <attribute name='keywords' />
                                    <attribute name='statuscodename' />
                                    <attribute name='createdbyyominame' />
                                    <attribute name='modifiedbyname' />
                                    <attribute name='versionnumber' />
                                    <attribute name='modifiedby' />
                                    <attribute name='createdby' />
                                    <attribute name='articlexml' />
                                    <attribute name='organizationidname' />
                                    <attribute name='exchangerate' />
                                    <attribute name='kbarticleid' />
                                    <attribute name='modifiedon' />
                                    <attribute name='title' />
                                    <attribute name='subjectidname' />
                                    <attribute name='modifiedonbehalfbyyominame' />
                                    <attribute name='createdbyname' />
                                    <attribute name='content' />
                                    <attribute name='number' />
                                    <attribute name='description' />
                                    <attribute name='modifiedonbehalfbyname' />
                                    <attribute name='createdonbehalfbyname' />
                                    <attribute name='overriddencreatedon' />
                                    <attribute name='entityimage_url' />
                                    <filter>
                                      <condition attribute='createdon' operator='this-week' />
                                    </filter>
                                  </entity>
                                </fetch>";
            collection = serviceSource.RetrieveMultiple(new FetchExpression(fetchXML));
            return collection;
        }


        /// <summary>
        /// Create Price List Records
        /// </summary>
        /// <param name="priceListColl"></param>
        /// <param name="serviceTarget"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        private static bool UnitPriceListRecords(EntityCollection priceListColl, IOrganizationService serviceTarget, TraceWriter log)
        {
            log.Info("Creating Price List records.");
            Entity pricelevel = new Entity("pricelevel");
            foreach (Entity entObj in priceListColl.Entities)
            {
                pricelevel.Id = new Guid(entObj.Attributes["pricelevelid"].ToString());
                if (entObj.Attributes.Contains("name"))
                {
                    pricelevel.Attributes["name"] = entObj.Attributes["name"].ToString();
                }

                serviceTarget.Create(pricelevel);
            }

            return true;
        }

        /// <summary>
        /// Create Unit Group records
        /// </summary>
        /// <param name="unitGroupColl"></param>
        /// <param name="serviceTarget"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        private static bool UnitGroupRecords(EntityCollection unitGroupColl, IOrganizationService serviceTarget, TraceWriter log)
        {
            log.Info("Creating Unit Group records.");
            Entity uomschedule = new Entity("uomschedule");
            foreach (Entity entObj in unitGroupColl.Entities)
            {
                uomschedule.Id = new Guid(entObj.Attributes["uomscheduleid"].ToString());
                if (entObj.Attributes.Contains("name"))
                {
                    uomschedule.Attributes["name"] = entObj.Attributes["name"].ToString();
                }
                if (entObj.Attributes.Contains("baseuomname"))
                {
                    uomschedule.Attributes["baseuomname"] = entObj.Attributes["baseuomname"].ToString();
                }

                serviceTarget.Create(uomschedule);
            }

            return true;
        }


        /// <summary>
        /// Retrieve Price list records
        /// </summary>
        /// <param name="serviceSource"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        private static EntityCollection RetrievePriceListRecords(IOrganizationService serviceSource, TraceWriter log)
        {
            log.Info("Fetching Price List Records....");
            EntityCollection priceListCollection = null;

            string fetchXML = @"<fetch distinct='false' mapping='logical' output-format='xml-platform' version='1.0' >
                                  <entity name='pricelevel' >
                                    <attribute name='createdonbehalfbyyominame' />
                                    <attribute name='modifiedonbehalfby' />
                                    <attribute name='transactioncurrencyidname' />
                                    <attribute name='statecode' />
                                    <attribute name='freighttermscode' />
                                    <attribute name='enddate' />
                                    <attribute name='description' />
                                    <attribute name='statecodename' />
                                    <attribute name='begindate' />
                                    <attribute name='createdonbehalfby' />
                                    <attribute name='transactioncurrencyid' />
                                    <attribute name='name' />
                                    <attribute name='pricelevelid' />
                                    <attribute name='importsequencenumber' />
                                    <attribute name='organizationidname' />
                                    <attribute name='modifiedbyyominame' />
                                    <attribute name='shippingmethodcode' />
                                    <attribute name='paymentmethodcode' />
                                    <attribute name='utcconversiontimezonecode' />
                                    <attribute name='createdbyyominame' />
                                    <attribute name='modifiedbyname' />
                                    <attribute name='versionnumber' />
                                    <attribute name='modifiedby' />
                                    <attribute name='createdby' />
                                    <attribute name='timezoneruleversionnumber' />
                                    <attribute name='shippingmethodcodename' />
                                    <attribute name='statuscodename' />
                                    <attribute name='paymentmethodcodename' />
                                    <attribute name='freighttermscodename' />
                                    <attribute name='modifiedon' />
                                    <attribute name='exchangerate' />
                                    <attribute name='modifiedonbehalfbyyominame' />
                                    <attribute name='statuscode' />
                                    <attribute name='createdbyname' />
                                    <attribute name='createdon' />
                                    <attribute name='organizationid' />
                                    <attribute name='createdonbehalfbyname' />
                                    <attribute name='modifiedonbehalfbyname' />
                                    <attribute name='overriddencreatedon' />
                                    <filter>
                                      <condition attribute='name' operator='like' value='%Zen%' />
                                    </filter>
                                  </entity>
                                </fetch>";
            priceListCollection = serviceSource.RetrieveMultiple(new FetchExpression(fetchXML));
            return priceListCollection;
        }

        /// <summary>
        /// Retrieve Unit Records
        /// </summary>
        /// <param name="serviceSource"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        private static EntityCollection RetrieveUnitRecords(IOrganizationService serviceSource, TraceWriter log)
        {
            log.Info("Fetching Unit Records....");
            EntityCollection unitCollection = null;

            string fetchXML = @"<fetch distinct='false' mapping='logical' output-format='xml-platform' version='1.0'>
                                  <entity name='uom' >
                                    <attribute name='uomscheduleid' />
                                    <attribute name='createdonbehalfbyyominame' />
                                    <attribute name='modifiedonbehalfby' />
                                    <attribute name='baseuomname' />
                                    <attribute name='uomscheduleidname' />
                                    <attribute name='createdonbehalfby' />
                                    <attribute name='overriddencreatedon' />
                                    <attribute name='modifiedbyexternalparty' />
                                    <attribute name='name' />
                                    <attribute name='modifiedbyexternalpartyname' />
                                    <attribute name='importsequencenumber' />
                                    <attribute name='organizationid' />
                                    <attribute name='modifiedbyyominame' />
                                    <attribute name='createdbyexternalparty' />
                                    <attribute name='createdbyexternalpartyname' />
                                    <attribute name='utcconversiontimezonecode' />
                                    <attribute name='createdbyyominame' />
                                    <attribute name='baseuom' />
                                    <attribute name='quantity' />
                                    <attribute name='isschedulebaseuom' />
                                    <attribute name='modifiedby' />
                                    <attribute name='createdby' />
                                    <attribute name='timezoneruleversionnumber' />
                                    <attribute name='modifiedbyname' />
                                    <attribute name='createdbyexternalpartyyominame' />
                                    <attribute name='modifiedon' />
                                    <attribute name='modifiedbyexternalpartyyominame' />
                                    <attribute name='modifiedonbehalfbyyominame' />
                                    <attribute name='createdbyname' />
                                    <attribute name='createdon' />
                                    <attribute name='createdonbehalfbyname' />
                                    <attribute name='uomid' />
                                    <attribute name='modifiedonbehalfbyname' />
                                    <attribute name='versionnumber' />
                                    <attribute name='isschedulebaseuomname' />
                                    <filter>
                                      <condition attribute='name' operator='like' value='%Zen%' uiname='Harshal Shinde' uitype='systemuser' />
                                    </filter>
                                  </entity>
                                </fetch>";
            unitCollection = serviceSource.RetrieveMultiple(new FetchExpression(fetchXML));
            return unitCollection;
        }

        /// <summary>
        /// Fetch Unit Group Records
        /// </summary>
        /// <param name="serviceSource"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        private static EntityCollection RetrieveUnitGroupRecords(IOrganizationService serviceSource, TraceWriter log)
        {
            log.Info("Fetching Unit Group Records....");
            EntityCollection unitGroupCollection = null;
            //Name contains Zen
            string fetchXML = @"<fetch top='50' >
                                  <entity name='uomschedule' >
                                    <attribute name='createdonbehalfbyyominame' />
                                    <attribute name='modifiedonbehalfby' />
                                    <attribute name='statecode' />
                                    <attribute name='baseuomname' />
                                    <attribute name='description' />
                                    <attribute name='statecodename' />
                                    <attribute name='createdonbehalfby' />
                                    <attribute name='modifiedbyexternalparty' />
                                    <attribute name='name' />
                                    <attribute name='modifiedbyexternalpartyname' />
                                    <attribute name='importsequencenumber' />
                                    <attribute name='organizationidname' />
                                    <attribute name='modifiedbyyominame' />
                                    <attribute name='createdbyexternalparty' />
                                    <attribute name='createdbyexternalpartyname' />
                                    <attribute name='utcconversiontimezonecode' />
                                    <attribute name='createdbyyominame' />
                                    <attribute name='timezoneruleversionnumber' />
                                    <attribute name='versionnumber' />
                                    <attribute name='modifiedby' />
                                    <attribute name='createdby' />
                                    <attribute name='uomscheduleid' />
                                    <attribute name='modifiedbyname' />
                                    <attribute name='statuscodename' />
                                    <attribute name='createdbyexternalpartyyominame' />
                                    <attribute name='modifiedon' />
                                    <attribute name='modifiedbyexternalpartyyominame' />
                                    <attribute name='modifiedonbehalfbyyominame' />
                                    <attribute name='statuscode' />
                                    <attribute name='createdbyname' />
                                    <attribute name='createdon' />
                                    <attribute name='organizationid' />
                                    <attribute name='createdonbehalfbyname' />
                                    <attribute name='modifiedonbehalfbyname' />
                                    <attribute name='overriddencreatedon' />
                                    <filter>
                                      <condition attribute='name' operator='like' value='%zen%' />
                                    </filter>
                                  </entity>
                                </fetch>";
            unitGroupCollection = serviceSource.RetrieveMultiple(new FetchExpression(fetchXML));
            return unitGroupCollection;
        }

        /// <summary>
        /// Create the products
        /// </summary>
        /// <param name="productColl"></param>
        /// <param name="serviceTarget"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        private static bool CreateProducts(EntityCollection productColl, IOrganizationService serviceTarget, TraceWriter log)
        {
            log.Info("Creating Product records.");
            Entity product = new Entity("product");
            foreach (Entity entObj in productColl.Entities)
            {
                if (entObj.Attributes.Contains("name"))
                {
                    product.Attributes["name"] = entObj.Attributes["name"].ToString();
                }
                if (entObj.Attributes.Contains("defaultuomscheduleid"))
                {
                    EntityReference entref = (EntityReference)entObj.Attributes["defaultuomscheduleid"];
                    var LookupId = entref.Id;
                    product.Attributes["defaultuomscheduleid"] = new EntityReference(entref.LogicalName, LookupId);
                }
                if (entObj.Attributes.Contains("defaultuomid"))
                {
                    EntityReference entref = (EntityReference)entObj.Attributes["defaultuomid"];
                    var LookupId = entref.Id;
                    product.Attributes["defaultuomid"] = new EntityReference(entref.LogicalName, LookupId);
                }


                serviceTarget.Create(product);
            }



            return true;
        }

        /// <summary>
        /// Fetch records from Product entity.
        /// </summary>
        /// <param name="service"></param>
        /// <returns>productCollection</returns>
        private static EntityCollection RetrieveProductRecords(IOrganizationService service, TraceWriter log)
        {
            log.Info("Fetching Records....");
            EntityCollection productCollection = null;

            string fetchXML = @"<?xml version='1.0'?>
                                    <fetch distinct='false' mapping='logical' output-format='xml-platform' version='1.0'>
                                    <entity name='product'>
                                    <attribute name='name'/>
                                    <attribute name='parentproductid'/>
                                    <attribute name='hierarchypath'/>
                                    <attribute name='validfromdate'/>
                                    <attribute name='validtodate'/>
                                    <attribute name='productstructure'/>
                                    <attribute name='iskit'/>
                                    <attribute name='productnumber'/>
                                    <attribute name='statecode'/>
                                    <attribute name='vendorpartnumber'/>
                                    <attribute name='vendorid'/>
                                    <attribute name='vendorname'/>
                                    <attribute name='producturl'/>
                                    <attribute name='defaultuomscheduleid'/>
                                    <attribute name='suppliername'/>
                                    <attribute name='subjectid'/>
                                    <attribute name='stockweight'/>
                                    <attribute name='stockvolume'/>
                                    <attribute name='isstockitem'/>
                                    <attribute name='statuscode'/>
                                    <attribute name='standardcost_base'/>
                                    <attribute name='standardcost'/>
                                    <attribute name='size'/>
                                    <attribute name='overriddencreatedon'/>
                                    <attribute name='quantityonhand'/>
                                    <attribute name='producttypecode'/>
                                    <attribute name='processid'/>
                                    <attribute name='modifiedon'/>
                                    <attribute name='modifiedbyexternalparty'/>
                                    <attribute name='modifiedonbehalfby'/>
                                    <attribute name='modifiedby'/>
                                    <attribute name='price_base'/>
                                    <attribute name='price'/>
                                    <attribute name='isreparented'/>
                                    <attribute name='iskit'/>
                                    <attribute name='exchangerate'/>
                                    <attribute name='description'/>
                                    <attribute name='defaultuomid'/>
                                    <attribute name='pricelevelid'/>
                                    <attribute name='quantitydecimal'/>
                                    <attribute name='currentcost_base'/>
                                    <attribute name='currentcost'/>
                                    <attribute name='transactioncurrencyid'/>
                                    <attribute name='createdon'/>
                                    <attribute name='createdbyexternalparty'/>
                                    <attribute name='createdonbehalfby'/>
                                    <attribute name='createdby'/>
                                    <attribute name='stageid'/>
                                    <order descending='false' attribute='hierarchypath'/>
                                    <filter type='and'>
                                    <condition attribute='statecode' operator='eq' value='0'/>
                                    </filter>
                                    </entity>
                                 </fetch>";

            productCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));

            return productCollection;
        }

        /// <summary>
        /// Create connection with D365 environment
        /// </summary>
        /// <param name="log"></param>
        /// <returns>Iorganization Service</returns>
        private static IOrganizationService Connection(TraceWriter log, string URL, string userName, string password)
        {
            log.Info("Connecting D365 '" + URL + "' environment...");
            IOrganizationService service = null;

            try
            {
                ClientCredentials clientCredentials = new ClientCredentials();
                clientCredentials.UserName.UserName = userName;
                clientCredentials.UserName.Password = password;


                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                OrganizationServiceProxy ornaziationServiceProxy = new OrganizationServiceProxy(new Uri("" + URL + "/XRMServices/2011/Organization.svc"), null, clientCredentials, null);

                ornaziationServiceProxy.Timeout = new TimeSpan(0, 10, 0);

                service = ornaziationServiceProxy;

                if (service != null)
                {
                    Guid userid = ((WhoAmIResponse)service.Execute(new WhoAmIRequest())).UserId;

                    if (userid != Guid.Empty)
                    {
                        log.Info("   Established Successfully...");
                    }
                }
                else
                {
                    log.Error("Failed to Established Connection!!!");
                }
            }
            catch (Exception ex)
            {
                log.Error("Exception caught - " + ex.Message);
            }

            return service;
        }

        public static IOrganizationService GetOrganizationServiceClientSecret(string clientId, string clientSecret, string organizationUri)
        {
            try
            {
                var conn = new CrmServiceClient($@"AuthType=ClientSecret;url={organizationUri};ClientId={clientId};ClientSecret={clientSecret}");

                return conn.OrganizationWebProxyClient != null ? conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
            }
            catch (Exception ex)
            {

                // Console.WriteLine("Error while connecting to CRM " + ex.Message);
                //log.Info("Creating Product records.");
                Console.ReadKey();
                return null;
            }
        }

    }
}
