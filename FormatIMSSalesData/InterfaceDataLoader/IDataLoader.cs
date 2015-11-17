using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InterfaceDataLoader
{
    public interface IDataLoader
    {
        string Name { get; }
        bool ImportData(Object Configuration, String DBConnectionString);
        

    }
}
