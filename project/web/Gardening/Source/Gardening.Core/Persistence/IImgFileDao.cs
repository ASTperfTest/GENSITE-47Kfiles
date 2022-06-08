using System.Collections;
using Gardening.Core.Domain;

namespace Gardening.Core.Persistence
{
    public interface IImgFileDao
    {
        ImgFile Get(string fileId);
        IList GetByParent(string parentId); //目前有topicId,parentId
        ImgFile Create(ImgFile file);
        ImgFile Update(ImgFile file);
        void Delete(string fileId);
        void DeleteByParent(string parentId);
    }
}
