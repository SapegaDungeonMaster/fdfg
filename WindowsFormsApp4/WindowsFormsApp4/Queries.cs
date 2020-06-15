using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp4
{
    class Queries
    {
        public static string selectgrup = "SELECT код_группы, количество_учащихся, название_специальности, группа FROM Группа, Специальность WHERE Группа.код_специальности=Специальность.код_специальности ";
        public static string selectdis = "SELECT * FROM Дисциплины";
        public static string selectdis_grup = "SELECT код_дисциплины_группы, название_дисциплины, группа, номер_кабинета,количество_часов_дисциплины FROM Дисциплины_группы, Дисциплины, Группа, Кабинет WHERE Дисциплины_группы.код_дисциплины=Дисциплины.код_дисциплины and Дисциплины_группы.код_группы=Группа.код_группы and Дисциплины_группы.код_кабинета=Кабинет.код_кабинета";
        public static string selectzaniat = "SELECT код_занятия, номер_кабинета, группа, дата_расписания,ФИО,название_дисциплины FROM Занятие, Кабинет, Расписание,Дисциплины,Преподаватель WHERE Занятие.код_кабинета=Кабинет.код_кабинета and Занятие.id_расписания=Расписание.id_расписания and Занятие.код_дисциплины=Дисциплины.код_дисциплины and Занятие.код_преподавателя=Преподаватель.код_преподавателя";
        public static string selectkabinet = "SELECT * FROM Кабинет";
        public static string selectprepod = "SELECT * FROM Преподаватель";
        public static string selectraspisan = "SELECT Расписание.id_расписания, Расписание.дата_расписания, Кабинет.номер_кабинета, Группа.группа, Дисциплины.название_дисциплины FROM Кабинет INNER JOIN(Дисциплины INNER JOIN (Группа INNER JOIN Расписание ON Группа.код_группы = Расписание.код_группы) ON Дисциплины.код_дисциплины = Расписание.код_дисципины) ON Кабинет.код_кабинета = Расписание.код_кабинета;";
        public static string selectspecial = "SELECT * FROM Специальность";
        public static string selectycha = "SELECT Учащиеся.код_учащегося, Учащиеся.ФИО, Учащиеся.телефон, Учащиеся.улица, Учащиеся.дата_рождения, Группа.группа FROM Группа INNER JOIN Учащиеся ON Группа.код_группы = Учащиеся.код_группы";
    }
}
