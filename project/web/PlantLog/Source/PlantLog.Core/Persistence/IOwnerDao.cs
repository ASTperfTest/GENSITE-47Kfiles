using System.Collections;
using PlantLog.Core.Domain;

namespace PlantLog.Core.Persistence
{
    public interface IOwnerDao
    {
        Owner Get(string ownerId);
        IList GetAll();
        Owner Create(Owner owner);
        Owner Update(Owner owner);
        void Delete(string ownerId);
    }
}
