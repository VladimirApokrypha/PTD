using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FuncionalPTD.FunctionalInterfaces.Behaviors
{
    /// <summary>
    /// интерфейс поведения добавления файла
    /// соотношения генподрядчика и субподрядчиков
    /// в древо каталогов
    /// </summary>
    public interface AddCASFileBehavior
    {
        /// <summary>
        /// метод, добавляющий  файл соотношения генподрядчика и субподрядчика
        /// </summary>
        void AddCASFile(string path);
    }
}
