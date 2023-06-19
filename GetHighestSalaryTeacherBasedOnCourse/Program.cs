using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetHighestSalaryTeacherBasedOnCourse
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Link Entity......");
            try
            {
                string conn = getConnectionString();

                Console.WriteLine("Connection Established......");

                CrmServiceClient service = new CrmServiceClient(conn);

                if (service.IsReady)
                {
                    var run = true;
                    while (run)
                    {
                        Console.WriteLine("Select Entity to Query:");
                        Console.WriteLine("1. Get Teacher List Who Has Highest Salary :");
                        Console.WriteLine("2. Get Total Amount in Order & related Quote: ");
                        Console.WriteLine("3. Retrieve All Accounts Along With Their Associated Contacts: ");
                        Console.WriteLine("4. Get Won & Lost Opportunity Associated With Contacts: ");
                        Console.WriteLine("5. Generate Quote From Opportunity");
                        Console.WriteLine("6. ");
                        Console.WriteLine("7. Exit");

                        string choice = Console.ReadLine();

                        switch (choice)
                        {
                            case "1":
                                getTeacherListFromCourseWhichHasHighestSalary(service);
                                break;

                            case "2":
                                getTotalAmountOrderAndRelatedQuote(service);
                                break;

                            case "3":
                                retrieveAccountsWithAssociatedContacts(service);
                                break;

                            case "4":
                                getClosedOpportunity(service);
                                break;

                            case "5":
                                generateQuoteFromOpportunity(service);
                                break;

                            case "6":
                                break;

                            case "7":
                                run = false;
                                break;

                            default:
                                Console.WriteLine("Invalid choice.");
                                break;
                        }
                    }
                    
                }
                else
                {
                    Console.WriteLine("Failed to connect with dynamics 365 crm");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error Message : " + e.Message);
            }
        }

        public static string getConnectionString()
        {
            string sUserKey = getAppSettingKey("UserKey");
            string sUserPassword = getAppSettingKey("UserPassword");
            string sEnvironment = getAppSettingKey("Environment");

            return $@" Url = {sEnvironment};AuthType = OAuth;UserName = {sUserKey}; Password = {sUserPassword};AppId = 51f81489-12ee-4a9e-aaae-a2591f45987d;RedirectUri = app://58145B91-0C36-4500-8554-080854F2AC97;LoginPrompt=Auto; RequireNewInstance = True";
        }

        public static string getAppSettingKey(string key)
        {
            return System.Configuration.ConfigurationManager.AppSettings[key].ToString();
        }

        public static void getTotalAmountOrderAndRelatedQuote(CrmServiceClient service)
        {
            QueryExpression query = new QueryExpression("salesorder");

            query.ColumnSet = new ColumnSet("name", "totalamount");

            LinkEntity quoteLink = new LinkEntity("salesorder", "quote", "quoteid", "quoteid", JoinOperator.Inner);
            quoteLink.Columns = new ColumnSet("name", "totalamount");
            quoteLink.EntityAlias = "quote";

            query.LinkEntities.Add(quoteLink);

            EntityCollection collection = service.RetrieveMultiple(query);

            foreach (Entity order in collection.Entities)
            {
                Console.WriteLine("Order Name: " + order.GetAttributeValue<string>("name"));
                Console.WriteLine("Order Total Amount: " + order.GetAttributeValue<Money>("totalamount").Value);
                Console.WriteLine("----------------------------------------------");

                Console.WriteLine("Quote Name: " + order.GetAttributeValue<AliasedValue>("quote.name").Value);
                string formattedamt = (((Money)order.GetAttributeValue<AliasedValue>("quote.totalamount").Value).Value).ToString("0.00");
                Console.WriteLine("Quote Total Amount: " + formattedamt);
                Console.WriteLine("----------------------------------------------");
            }
        }

        public static void getTeacherListFromCourseWhichHasHighestSalary(CrmServiceClient service)
        {
            QueryExpression courseQuery = new QueryExpression("new_course");
            courseQuery.ColumnSet = new ColumnSet("new_courseid", "new_coursename");

            LinkEntity teacherLink = new LinkEntity("new_course", "new_teacher", "new_courseid", "cr7fa_relatedcourse", JoinOperator.Inner);
            teacherLink.Columns = new ColumnSet("new_teachername", "cr7fa_teachersalary", "cr7fa_relatedcourse");
            teacherLink.EntityAlias = "Teacher";

            courseQuery.LinkEntities.Add(teacherLink);

            EntityCollection collection = service.RetrieveMultiple(courseQuery);

            var teachersByCourse = collection.Entities
                .Where(t => t.Contains("Teacher.cr7fa_relatedcourse") && t.Contains("Teacher.cr7fa_teachersalary"))
                .GroupBy(t => ((EntityReference)(t.GetAttributeValue<AliasedValue>("Teacher.cr7fa_relatedcourse").Value)).Id)
                .Select(g => g.OrderByDescending(t => ((Money)(t.GetAttributeValue<AliasedValue>("Teacher.cr7fa_teachersalary").Value)).Value).FirstOrDefault());

            Console.WriteLine("Teacher List with Highest Salary per Course:");
            Console.WriteLine("----------------------------------------------");

            foreach (Entity teacher in teachersByCourse)
            {
                string teacherName = teacher.GetAttributeValue<AliasedValue>("Teacher.new_teachername").Value.ToString();
                decimal value = ((Money)teacher.GetAttributeValue<AliasedValue>("Teacher.cr7fa_teachersalary").Value).Value;
                Money salary = new Money(value);
                Guid courseId = ((EntityReference)(teacher.GetAttributeValue<AliasedValue>("Teacher.cr7fa_relatedcourse").Value)).Id;

                string courseName = collection.Entities
                    .Where(c => c.Id == courseId)
                    .Select(c => c.GetAttributeValue<string>("new_coursename"))
                    .FirstOrDefault();

                Console.WriteLine($"Teacher: {teacherName} \nCourse: {courseName} \nSalary: {salary.Value.ToString("0.00")}\n");
            }
        }

        public static void retrieveAccountsWithAssociatedContacts(CrmServiceClient service)
        {
            QueryExpression query = new QueryExpression("account");
            query.ColumnSet = new ColumnSet("name");

            LinkEntity contactLink = new LinkEntity("account", "contact", "accountid", "parentcustomerid", JoinOperator.Inner);
            contactLink.Columns = new ColumnSet("fullname");
            contactLink.EntityAlias = "contactAlias";
            query.LinkEntities.Add(contactLink);
            EntityCollection collection = service.RetrieveMultiple(query);

            foreach (Entity account in collection.Entities)
            {
                Console.WriteLine("Account: " + account.GetAttributeValue<string>("name"));

                if (account.Contains("contactAlias.fullname"))
                {
                    AliasedValue aliasedValue = account["contactAlias.fullname"] as AliasedValue;
                    if (aliasedValue != null)
                    {
                        string contactFullName = aliasedValue.Value.ToString();
                        Console.WriteLine("Contact: " + contactFullName);
                    }
                    Console.WriteLine("-----------------------------------");
                }
            }
        }

        public static void getClosedOpportunity(CrmServiceClient service)
        {
            QueryExpression query = new QueryExpression("account");
            query.ColumnSet = new ColumnSet("name");

            LinkEntity contactLink = new LinkEntity("account", "contact", "accountid", "parentcustomerid", JoinOperator.Inner);
            contactLink.Columns = new ColumnSet("fullname");
            contactLink.EntityAlias = "contactAlias";
            query.LinkEntities.Add(contactLink);

            LinkEntity opportunityLink = new LinkEntity("contact", "opportunity", "contactid", "parentcontactid", JoinOperator.Inner);
            opportunityLink.Columns = new ColumnSet("name");

            opportunityLink.LinkCriteria = new FilterExpression();
            opportunityLink.LinkCriteria.FilterOperator = LogicalOperator.Or;

            ConditionExpression closedCondition = new ConditionExpression("statecode", ConditionOperator.Equal, 2);
            opportunityLink.LinkCriteria.AddCondition(closedCondition);
            ConditionExpression wonCondition = new ConditionExpression("statuscode", ConditionOperator.Equal, 1);
            opportunityLink.LinkCriteria.AddCondition(wonCondition);

            opportunityLink.EntityAlias = "opportunityAlias";
            contactLink.LinkEntities.Add(opportunityLink);

            EntityCollection collection = service.RetrieveMultiple(query);
            List<Entity> entities = collection.Entities.ToList();

            var contactByAccount = entities
                .GroupBy(t => t.GetAttributeValue<string>("name"));

            foreach (var contact in contactByAccount)
            {
                string accountName = contact.Key;

                var opportunityByContact = contact
                .GroupBy(t => t.GetAttributeValue<AliasedValue>("contactAlias.fullname").Value.ToString());

                string AcCheck = null;

                foreach (var item in opportunityByContact)
                {
                    string contactName = item.Key;
                    string contactCheck = null;
                    int opportunityCount = 0;
                    foreach (var entity in item)
                    {
                        if (AcCheck == null)
                        {
                            Console.WriteLine("Account Name: " + accountName);
                            AcCheck = accountName;
                        }
                        if (contactCheck == null)
                        {
                            Console.WriteLine("Contact Name: " + contactName);
                            contactCheck = contactName;
                        }
                        if (entity.Contains("opportunityAlias.name"))
                        {
                            AliasedValue aliasedValue = entity["opportunityAlias.name"] as AliasedValue;

                            if (aliasedValue != null)
                            {
                                opportunityCount++;
                                string opportunityName = aliasedValue.Value.ToString();
                                Console.WriteLine("Opportunity Name: " + opportunityName);
                            }
                        }
                    }
                    Console.WriteLine("Total Opportunity: " + opportunityCount);
                }
                Console.WriteLine("-----------------------------------------");
            }
        }

        public static void generateQuoteFromOpportunity(CrmServiceClient service)
        {
            QueryExpression query = new QueryExpression("opportunity");

            ColumnSet columnSet = new ColumnSet("opportunityid", "name");
            query.ColumnSet = columnSet;

            query.PageInfo = new PagingInfo { Count = 20, PageNumber = 1 };

            EntityCollection collection = service.RetrieveMultiple(query);

            foreach (Entity opportunity in collection.Entities)
            {
                Guid oppID = opportunity.GetAttributeValue<Guid>("opportunityid");
                string name = opportunity.GetAttributeValue<string>("name");

                Console.WriteLine("Opportunity Id: " + oppID + "\n\tName: " + name);
                Console.WriteLine("---------------------------------------------------");
            }

            Console.WriteLine("Enter Opportunity Id You Want To Generate Quote: ");
            var generateQuoteId = Console.ReadLine();

            Guid generateQuoteGuid = new Guid(generateQuoteId);
            GenerateQuoteFromOpportunityRequest generateQuoteRequest = new GenerateQuoteFromOpportunityRequest
            {
                ColumnSet = new ColumnSet(true),
                OpportunityId = generateQuoteGuid
            };

            GenerateQuoteFromOpportunityResponse quoteResponse =
                        (GenerateQuoteFromOpportunityResponse)service.Execute(generateQuoteRequest);

            Guid quoteId = quoteResponse.Entity.Id;

            Console.WriteLine("Quote Created Successfully..........");
            Console.ReadLine();
        }
    }
}