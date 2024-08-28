using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace ImportExcel
{
    public abstract class BaseModel
    {
        public Type GetType ()
        {
            return this.GetType();
        }
    }
}
