using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IT_School
{
    internal class Organization
    {
        public string AccName { get; set; } //Наименование аккредитационного органа
        public string Name { get; set; } //Полное наименование, местонахождение организации
        public string Adress { get; set; } //Местонахождение юридического лица
        public string GeoData { get; set; } //Геоданные
        public string WorkTime { get; set; } //Режим работы
        public string GosID { get; set; } //Государственный регистрационный номер записи о создании юридического лица
        public string Inn { get; set; } //Идентификационный номер налогоплательщика-организации
        public string DateBegin { get; set; } //Дата принятия решения о государственной аккредитации
        public string GosAccReq { get; set; } //Реквизиты свидетельств
        public string DateExpire { get; set; } //Срок окончания свидетельства
        public string EduSpecs { get; set; } //Перечень аккредитованных образовательных программ, укрупненных групп направлений подготовки и специальностей
        public string ReMake { get; set; } //Основание и дата переоформления свидетельства о государственной аккредитации
        public string StopStart { get; set; } //Основание и даты приостановлении и  возобновлении действия свидетельства о государственной аккредитации
        public string StopExec { get; set; } //Основание и дата лишения свидетельства о государственной аккредитации 
        public string Stop { get; set; } //Основание и дата прекращения действия  свидетельства о государственной аккредитации 
    }
}
