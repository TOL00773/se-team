using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IT_School
{
    internal class Organization
    {
        string AccName { get; set; } //Наименование аккредитационного органа
        string Name { get; set; } //Полное наименование, местонахождение организации
        string Adress { get; set; } //Местонахождение юридического лица
        string GeoData { get; set; } //Геоданные
        string WorkTime { get; set; } //Режим работы
        string GosID { get; set; } //Государственный регистрационный номер записи о создании юридического лица
        string Inn { get; set; } //Идентификационный номер налогоплательщика-организации
        string DateBegin { get; set; } //Дата принятия решения о государственной аккредитации
        string GosAccReq { get; set; } //Реквизиты свидетельств
        string DateExpire { get; set; } //Срок окончания свидетельства
        List<string> EduSpecs { get; set; } //Перечень аккредитованных образовательных программ, укрупненных групп направлений подготовки и специальностей
        string ReMake { get; set; } //Основание и дата переоформления свидетельства о государственной аккредитации
        string StopStart { get; set; } //Основание и даты приостановлении и  возобновлении действия свидетельства о государственной аккредитации
        string StopExec { get; set; } //Основание и дата лишения свидетельства о государственной аккредитации 
        string Stop { get; set; } //Основание и дата прекращения действия  свидетельства о государственной аккредитации 
    }
}
