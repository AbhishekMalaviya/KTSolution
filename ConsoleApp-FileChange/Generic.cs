using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp_FileChange
{
    internal class Generic<T,U>
    {
        public virtual void WriteData(T t, U u) 
        {
            Console.WriteLine($"There are type value: {t} and {u}"); 
        }
    }

    class Generic1<T> where T : IRequest
    {
        public void Method1<U>(T t,U u)
        {
            Console.WriteLine($"Generic response:{t.GetSecret("key1")}. U type: {u.GetType()}"); 
        }
    }

    interface IRequest
    {
        string GetSecret(string key);
    }

    internal class Request : IRequest
    {
        public string TenantId { get; set;}

        public Request(string tenantId)
        {
            TenantId = tenantId;
        }

        public string GetSecret(string key)
        {
            return $"secret {key}";
        }
    }
}
