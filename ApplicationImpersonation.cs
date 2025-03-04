using System;
using System.DirectoryServices;
using System.Runtime.InteropServices;
using static Vanara.PInvoke.ActiveDS;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 3)
        {
            Console.WriteLine("Usage: Program <DomainName> <DomainControllerIP> <UserName>");
            return;
        }

        string domainName = args[0];
        string dcIp = args[1];
        string userName = args[2];

        string domainBaseDn = GetDomainBaseDn(domainName);
        string configBaseDn = $"CN=Microsoft Exchange,CN=Services,CN=Configuration,{domainBaseDn}";
        string organizationDn = GetOrganizationContainer(dcIp, configBaseDn);

        if (organizationDn == null)
        {
            Console.WriteLine("Could not find the organization container.");
            return;
        }

        string rbacDn = $"CN=RBAC,{organizationDn}";

        if (!HasAccessToRbac(dcIp, rbacDn))
        {
            Console.WriteLine("Insufficient permissions to access the RBAC container.");
            return;
        }

        string roleName = "ApplicationImpersonation";
        string roleDn = $"CN={roleName},CN=Roles,{rbacDn}";
        string roleAssignmentBaseDn = $"CN=Role Assignments,{rbacDn}";

        Console.WriteLine($"Domain Base DN: {domainBaseDn}");
        Console.WriteLine($"Role DN: {roleDn}");
        Console.WriteLine($"Role Assignment Base DN: {roleAssignmentBaseDn}");

        try
        {
            string ldapPath = $"LDAP://{dcIp}";
            string userDn = GetUserDn(userName, ldapPath);

            if (userDn == null)
            {
                Console.WriteLine($"User {userName} not found.");
                return;
            }

            Console.WriteLine($"User DN found: {userDn}");

            // 动态获取 msExchVersion
            object msExchVersionValue = GetMsExchVersionFromExchangeServers(dcIp, domainBaseDn);
            if (msExchVersionValue == null)
            {
                Console.WriteLine("Could not determine msExchVersion.");
                return;
            }

            CreateOrUpdateRoleAssignment(dcIp, roleAssignmentBaseDn, $"CN=ApplicationImpersonation-{userName}", userDn, roleDn, msExchVersionValue);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    static string GetDomainBaseDn(string domainName)
    {
        var parts = domainName.Split('.');
        var dnParts = Array.ConvertAll(parts, part => $"DC={part}");
        return string.Join(",", dnParts);
    }

    static string GetUserDn(string userName, string ldapPath)
    {
        using (DirectoryEntry entry = new DirectoryEntry(ldapPath))
        {
            using (DirectorySearcher searcher = new DirectorySearcher(entry))
            {
                searcher.Filter = $"(sAMAccountName={userName})";
                searcher.PropertiesToLoad.Add("distinguishedName");

                Console.WriteLine($"Executing search with filter: {searcher.Filter}");
                SearchResult result = searcher.FindOne();

                if (result != null && result.Properties["distinguishedName"].Count > 0)
                {
                    return result.Properties["distinguishedName"][0].ToString();
                }
            }
        }
        return null;
    }

    static string GetOrganizationContainer(string dcIp, string configBaseDn)
    {
        using (DirectoryEntry configEntry = new DirectoryEntry($"LDAP://{dcIp}/{configBaseDn}"))
        {
            foreach (DirectoryEntry child in configEntry.Children)
            {
                if (child.SchemaClassName == "msExchOrganizationContainer")
                {
                    return child.Properties["distinguishedName"].Value.ToString();
                }
            }
        }
        return null;
    }

    static bool HasAccessToRbac(string dcIp, string rbacDn)
    {
        try
        {
            using (DirectoryEntry rbacEntry = new DirectoryEntry($"LDAP://{dcIp}/{rbacDn}"))
            {
                var schemaClassName = rbacEntry.SchemaClassName;
                return true;
            }
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
        catch (COMException)
        {
            return false;
        }
    }

    static object GetMsExchVersionFromExchangeServers(string dcIp, string domainBaseDn)
    {
        string exchangeServersDn = $"CN=Exchange Servers,OU=Microsoft Exchange Security Groups,{domainBaseDn}";
        using (DirectoryEntry exchangeServersEntry = new DirectoryEntry($"LDAP://{dcIp}/{exchangeServersDn}"))
        {
            if (exchangeServersEntry.Properties.Contains("msExchVersion"))
            {
                return exchangeServersEntry.Properties["msExchVersion"].Value;
            }
        }
        return null;
    }

    static void CreateOrUpdateRoleAssignment(string dcIp, string roleAssignmentBaseDn, string roleAssignmentCn, string userDn, string roleDn, object msExchVersionValue)
    {
        string roleAssignmentDn = $"{roleAssignmentCn},{roleAssignmentBaseDn}";
        Console.WriteLine($"Role Assignment DN: {roleAssignmentDn}");

        using (DirectoryEntry parentEntry = new DirectoryEntry($"LDAP://{dcIp}/{roleAssignmentBaseDn}"))
        {
            DirectorySearcher searcher = new DirectorySearcher(parentEntry)
            {
                Filter = $"(cn={roleAssignmentCn})"
            };

            SearchResult result = searcher.FindOne();
            DirectoryEntry roleAssignment;

            if (result != null)
            {
                Console.WriteLine("Role assignment already exists. Updating...");
                roleAssignment = result.GetDirectoryEntry();
            }
            else
            {
                Console.WriteLine("Creating new role assignment...");
                roleAssignment = parentEntry.Children.Add(roleAssignmentCn, "msExchRoleAssignment");
            }

            roleAssignment.Properties["msExchUserLink"].Value = userDn;
            roleAssignment.Properties["msExchRoleLink"].Value = roleDn;

            // 使用 unchecked 处理大于 int 的值
            roleAssignment.Properties["msExchRoleAssignmentFlags"].Value = CreateLargeInteger(unchecked((int)0x80010000), (int)0x02000002);

            roleAssignment.Properties["msExchVersion"].Value = msExchVersionValue;
            roleAssignment.Properties["systemFlags"].Value = 1073741824;

            Console.WriteLine("Saving changes to Active Directory...");
            roleAssignment.CommitChanges();

            Console.WriteLine("Role assignment processed successfully.");
        }
    }

    static IADsLargeInteger CreateLargeInteger(int highPart, int lowPart)
    {
        IADsLargeInteger largeInteger = (IADsLargeInteger)new LargeInteger();
        largeInteger.HighPart = highPart;
        largeInteger.LowPart = lowPart;
        return largeInteger;
    }
}
