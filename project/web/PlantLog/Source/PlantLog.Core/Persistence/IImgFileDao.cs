using System.Collections;
using PlantLog.Core.Domain;

namespace PlantLog.Core.Persistence
{
    public interface IImgFileDao
    {
        ImgFile Get(string fileId);
        IList GetByEntry(string entryId);
        ImgFile Create(ImgFile file);
        ImgFile Update(ImgFile file);
        void Delete(string fileId);
        void DeleteByEntry(string entryId);
    }
}
