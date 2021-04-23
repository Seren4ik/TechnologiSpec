using CodeStack.SwEx.AddIn;
using CodeStack.SwEx.AddIn.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System.IO;
using ExcelDataReader;
using System.Data;
using System.Collections;

namespace TechnologiSpec
{

    [ComVisible(true), Guid("815C9D76-3C11-4EAC-9B18-0EC714227DAB")]
    [AutoRegister("TechnologiSpec", "Add Technologic Specification")]

    public class Class1 : SwAddInEx
    {
        public class Component1
        {
            public IModelDoc2 sw;
            public int quantity;
            public string pos;
            public string designation;
            public string compname;
            public string config;

            public Component1(IModelDoc2 sw, int quantity, string pos, string designation, string compname, string config)
            {
                this.sw = sw;
                this.quantity = quantity;
                this.pos = pos;
                this.designation = designation;
                this.compname = compname;
                this.config = config;

            }

        }

        public class ComponentProp
        {
            public IModelDoc2 sw;
            public string nameConfig;
            public string property;


            public ComponentProp(IModelDoc2 sw, string nameConfig, string property)
            {
                this.sw = sw;
                this.nameConfig = nameConfig;
                this.property = property;

            }

            public string GetProp(IModelDoc2 sw, string nameConfig, string property)//Получение свойства
            {
                string prop1 = sw.GetCustomInfoValue(nameConfig, property);
                if (prop1 == "")
                {
                    string prop3 = sw.CustomInfo[property];
                    return prop3;
                }
                else
                {
                    return prop1;
                }
                
            }
            public double GetProp2(IModelDoc2 sw, string nameConfig, string property)//Получение свойства технологического
            {
                try
                {
                    string prop2 = sw.GetCustomInfoValue(nameConfig, property);
                    if (String.IsNullOrEmpty(prop2))
                    {
                        double x1 = 0;
                        return x1;
                    }
                    else
                    {
                        double.TryParse(prop2.Replace('.', ','), out double x1);
                        return x1;
                    }
                }
                catch (Exception)
                {

                    double x1 = 0;
                    return x1;
                }


            }
        }


        List<Component1> ObjectsList = new List<Component1>();
        ArrayList ListObject1 = new ArrayList();

        public object p;

        private enum Commands_e
        {
            AddTechnoSpec
        }

        public override bool OnConnect()
        {
            AddCommandGroup<Commands_e>(OnButtonClick);
            return true;
        }

        private void OnButtonClick(Commands_e cmd)
        {
            switch (cmd)
            {
                case Commands_e.AddTechnoSpec:
                    AddTechnoSpec1();
                    break;
            }
        }

        #region Формирование переменных (определение сборок и деталей)
        private void AddTechnoSpec1()
        {

            IModelDoc2 swModel;

            swModel = (IModelDoc2)App.ActiveDoc;//Экземпляр документа главной сборки
            string nameConf = swModel.IGetActiveConfiguration().Name.ToString();//Имя активной конфигурации
            string designation1 = swModel.GetCustomInfoValue(nameConf, "Обозначение"); //Обозначение главной сборки
            string compName = swModel.GetCustomInfoValue(nameConf, "Наименование"); //Обозначение главной сборки
            string pathcomp = swModel.GetPathName(); //Путь к файлу

            var assemblyComponent = new Component1(swModel, 1, "1", designation1, compName, nameConf);
            ObjectsList.Add(assemblyComponent); //Добавление в список  

            cicle(swModel, designation1, "1", 1, compName);//Запуск цикла
            AddExcelFile(designation1, pathcomp); //Создание файла ексель
            GetInfo(ObjectsList, pathcomp);// Определение свойств, передает словарь объектов с количеством
            ObjectsList.Clear();
            FillingExcel(ListObject1, pathcomp);// Заполнение екселя
            ListObject1.Clear();
            App.SendMsgToUser2("Файл создан", 0, 0);
        }

        private void cicle(IModelDoc2 swModel, string designation, string parentPosition, int position, string compname)
        {
  
            AssemblyDoc swAssembly;
            Object[] Components;

            if (swModel is AssemblyDoc)
            {
                swAssembly = (AssemblyDoc)swModel; //Экземпляр сборки
                Components = (Object[])swAssembly.GetComponents(true);//Список компонентов сборки
                List<Component1> childComponents = new List<Component1>();
                int lent1 = Components.Length;// Общее число компанентов

                foreach (Object obj in Components)
                {
                    Component2 obj1 = (Component2)obj;
                    IModelDoc2 component2 = (IModelDoc2)obj1.GetModelDoc();
                    //Проверка на погашен ли компонент
                    if (component2 == null)
                    {
                        
                    }
                    else
                    {
                        string nameConf2 = obj1.ReferencedConfiguration;
                        string nameConf3 = nameConf2.Replace('x', 'х').Replace('M', 'М').Replace('C', 'С').Replace('B', 'В').Replace('H', 'Н').Replace('A', 'А').Replace('T', 'Т').Replace('X', 'Х');
                        string componentDesignation = component2.GetCustomInfoValue(nameConf2, "Обозначение");
                        string componentName = component2.GetCustomInfoValue(nameConf2, "Наименование");
                        string type = component2.CustomInfo["IsFastener"];

                        if (type == "1")
                        {
                            if (childComponents.Any(c => c.compname == nameConf3))
                            {
                                var existingComponent = childComponents.First(c => c.compname == nameConf3);
                                existingComponent.quantity++;
                            }
                            else
                            {
                                string currentPosition = parentPosition + ". " + position.ToString();
                                position = position + 1;
                                var childComponent = new Component1(component2, 1, currentPosition, componentDesignation, nameConf3, nameConf2);
                                childComponents.Add(childComponent);

                            }

                        }
                        else
                        {
                            if (childComponents.Any(c => c.designation == componentDesignation) && childComponents.Any(c => c.compname == componentName))
                            {
                                var existingComponent = childComponents.First(c => c.compname == componentName && c.designation == componentDesignation);
                                existingComponent.quantity++;
                            }
                            else
                            {
                                string currentPosition = parentPosition + ". " + position.ToString();
                                position = position + 1;
                                var childComponent = new Component1(component2, 1, currentPosition, componentDesignation, componentName, nameConf2);
                                childComponents.Add(childComponent);
                                if (component2 is AssemblyDoc)
                                    cicle(component2, componentDesignation, currentPosition, 1, componentName);
                            }
                        }

                    }                

                }
                ObjectsList.AddRange(childComponents);
            }

            else
            {
                string nameConf1 = swModel.IGetActiveConfiguration().Name;
                ObjectsList.Add(new Component1(swModel, 1, parentPosition + ". 1", designation, compname, nameConf1));

            }

        }


        #endregion
        #region Чтение свойств Solid
        private void GetInfo(List<Component1> objects, string pathname)
        {

            foreach (Component1 component in objects.OrderBy(c => c.pos))
            {
                Dictionary<object, object> DictObject1 = new Dictionary<object, object>();

                string[] list1 = {"Резка (труба, круг, уголок)", "Рубка листовой оцинк. стали", "Резка листового материала лазером",
                    "Зачистка деталей", "Гибка деталей","Рихтовка деталей","Сверление","Зенкование отверстий","Нарезание резьбы",
                 "Точечная сварка",  "Вальцовка деталей", "Торцовка доски", "Калибровка доски", "Шлифование торцов", "Фрезерование радиусов",
                "Изготовление на станке с ЧПУ", "Сверление отверстий","Изготовление шип/паз", "Резка поликарбоната","Шлифование деревянной детали",
                    "Нанесение пропитки, лака, краски на деревянные детали", "Сушка деревянных деталей", "Разделка кромок под сварку",
                "Сборка деталей в кондукторе", "Сварка деталей", "Зачистка после сварки", "Зачистка металлических деталей перед покраской",
                "Покраска (промывка, сушка, нанесение порошка, установка в печь, съем деталей)", "Комплектация деталей на изделие","Сборка изделия",
                "Упаковка изделия", "Перемещение пустой формы",  "Перемещение заполненной формы", "Сварка каркаса", "Разборка пустой формы (2 раза)",
                "Смазка формы", "Сборка формы (2 раза)",  "Замес бетона", "Заливка бетона", "Установка закладных", "Выемка изделия из формы",  "Очистка формы",
                "Обработка фасок, удаление облоя", "Шлифование изделия", "Полирование изделия", "Погрузка изделия","Изготовление по кооперации","Прочие слесарные операции"
                };

                IModelDoc2 swModel = component.sw;
                int quantity1 = component.quantity;
                string position = component.pos;
                string nameConf = component.config;
           /*   string type1 = swModel.CustomInfo["Тип"];
                string type2 = swModel.GetCustomInfoValue(nameConf, "Тип");
                string sort1 = swModel.CustomInfo["Сортамент"];
                string sort2 = swModel.GetCustomInfoValue(nameConf, "Сортамент");*/
                string volume2 = swModel.GetCustomInfoValue(nameConf, "Масса");
                string designation = swModel.GetCustomInfoValue(nameConf, "Обозначение");
                string name1 = component.compname;

                var Comp1 = new ComponentProp(swModel, nameConf, "Толщина");
                double tolsh1 = Comp1.GetProp2(swModel, nameConf, "Толщина");

                var Comp2 = new ComponentProp(swModel, nameConf, "Ширина");
                double widht = Comp2.GetProp2(swModel, nameConf, "Ширина");

                var Comp3 = new ComponentProp(swModel, nameConf, "Длина");
                double length1 = Comp3.GetProp2(swModel, nameConf, "Длина");

                var Comp4 = new ComponentProp(swModel, nameConf, "Тип");
                string type1 = Comp4.GetProp(swModel, nameConf, "Тип");

                var Comp5 = new ComponentProp(swModel, nameConf, "Сортамент");
                string sort1 = Comp5.GetProp(swModel, nameConf, "Сортамент");

                double volume = (tolsh1 * widht * length1) / 1000000000 * quantity1 * 1.25;//Обьем м3 на количество
                double surface_area2 = ((tolsh1 * widht) + (tolsh1 * length1) + (length1 * widht)) / 1000000;//Площадь поверхности расчетная

                var Comp6 = new ComponentProp(swModel, nameConf, "Площадь поверхности");
                double surface_area = Comp6.GetProp2(swModel, nameConf, "Площадь поверхности");//Площадь поверхности из модели

                double propitka = surface_area * 0.14 * 3 * 1.15 * quantity1;//Пропитка
                double powder = surface_area * 0.14 * 1.15 * quantity1;//Порошок
                double powder1 = surface_area *0.55 * 0.14 * 1.15 * quantity1;//Порошок для трубы


                if (type1 == "Доска" || type1 == "Брус клеенный" || type1 == "Брусок" || type1 == "Пиловочник" || type1 == "Брус")
                {
                    DictObject1.Add(0, position);
                    DictObject1.Add(1, designation);
                    DictObject1.Add(2, name1);
                    DictObject1.Add(3, sort1);
                    DictObject1.Add(4, tolsh1);
                    DictObject1.Add(5, widht);
                    DictObject1.Add(6, length1);
                    DictObject1.Add(7, quantity1);
                    DictObject1.Add(8, Math.Round(volume, 4).ToString());
                    DictObject1.Add(9, " ");
                    DictObject1.Add(10, Math.Round(propitka, 3).ToString());
                    DictObject1.Add(11, " ");
                }
                
                else if(type1.Contains("Труба"))
                {
                    DictObject1.Add(0, position);
                    DictObject1.Add(1, designation);
                    DictObject1.Add(2, name1);
                    DictObject1.Add(3, sort1);
                    DictObject1.Add(4, tolsh1);
                    DictObject1.Add(5, widht);
                    DictObject1.Add(6, length1);
                    DictObject1.Add(7, quantity1);
                    DictObject1.Add(8, " ");
                    DictObject1.Add(9, volume2);
                    DictObject1.Add(10, " ");
                    DictObject1.Add(11, Math.Round(powder1, 3).ToString());
                }                   

                else
                {
                    DictObject1.Add(0, position);
                    DictObject1.Add(1, designation);
                    DictObject1.Add(2, name1);
                    DictObject1.Add(3, sort1);
                    DictObject1.Add(4, tolsh1);
                    DictObject1.Add(5, widht);
                    DictObject1.Add(6, length1);
                    DictObject1.Add(7, quantity1);
                    DictObject1.Add(8, " ");
                    DictObject1.Add(9, volume2);
                    DictObject1.Add(10, " ");
                    DictObject1.Add(11, Math.Round(powder, 3).ToString());
                }

                foreach (string propname in list1)
                {
                    var Comp = new ComponentProp(swModel, nameConf, propname);
                    DictObject1.Add(propname, Comp.GetProp2(swModel, nameConf, propname));
                }

                ListObject1.Add(DictObject1); //Добавление списка свойств компонента в главный список
            }

        }
        #endregion

        #region Создание файла ексель и шапки
        static void AddExcelFile(string nameproject, string pathname)
        {
            try
            {
                using (ExcelHelper helper = new ExcelHelper())
                {
                    if (helper.Open(filpath: pathname + ".xlsx"))
                    {
                        helper.Unite(Cell1: "A1", Cell2: "L1");
                        helper.Unite(Cell1: "A2", Cell2: "L2");
                        helper.Unite(Cell1: "A3", Cell2: "A4");
                        helper.Unite(Cell1: "B3", Cell2: "B4");
                        helper.Unite(Cell1: "C3", Cell2: "C4");
                        helper.Unite(Cell1: "D3", Cell2: "D3");
                        helper.Unite(Cell1: "E3", Cell2: "G3");
                        helper.Unite(Cell1: "H3", Cell2: "H4");
                        helper.Unite(Cell1: "I3", Cell2: "J3");
                        helper.Unite(Cell1: "K3", Cell2: "K4");
                        helper.Unite(Cell1: "L3", Cell2: "L4");
                        helper.Unite(Cell1: "M3", Cell2: "M4");
                        helper.Unite(Cell1: "N3", Cell2: "N4");
                        helper.Unite(Cell1: "O3", Cell2: "O4");

                        helper.Set(column: "A1", row: "L1", data: "Технологический маршрут изготовления (нормы времени)");
                        helper.Set(column: "A2", row: "L2", data: nameproject);
                        helper.Set(column: "A3", row: "A4", data: "№ п/п");
                        helper.Set(column: "B3", row: "B4", data: "Номер детали");
                        helper.Set(column: "C3", row: "C4", data: "Название");
                        helper.Set(column: "D3", row: "D3", data: "Сортамент");
                        helper.Set(column: "E3", row: "G3", data: "Размеры");
                        helper.Set(column: "I3", row: "J3", data: "Н/расхода");
                        helper.Set(column: "E4", row: "E4", data: "Толщ.");
                        helper.Set(column: "F4", row: "F4", data: "Шир.");
                        helper.Set(column: "G4", row: "G4", data: "Длин.");
                        helper.Set(column: "H3", row: "H4", data: "Kо-во");
                        helper.Set(column: "I4", row: "I4", data: "Объем м3");
                        helper.Set(column: "J3", row: "J4", data: "Вес/кг");
                        helper.Set(column: "K3", row: "K4", data: "Пропитка, кг");
                        helper.Set(column: "L3", row: "L4", data: "Краска порошковая, кг");
                        helper.Set(column: "M3", row: "M4", data: "Операции");
                        helper.Set(column: "N3", row: "N4", data: "Время (мин)");
                        helper.Set(column: "O3", row: "O4", data: "Время общ.(мин)");


                        helper.save();

                    }
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            {

            }

        }


        #endregion


        #region Заполнение екселя
        private void FillingExcel(ArrayList listname, string pathname)
        {
            try
            {
                using (ExcelHelper helper = new ExcelHelper())
                {
                    if (helper.Open(filpath: pathname + ".xlsx"))
                    {
                        int i, j, l;
                        j = 4;

                        foreach (Dictionary<object, object> list34 in listname)
                        {
                            l = 0;
                            i = 0;
                            j = j + 1;

                            foreach (KeyValuePair<object, object> value1 in list34)
                            {
                                l = l + 1;
                                i = i + 1;
                                if (l <= 11)
                                {

                                    helper.Cells(column: i, row: j, data: value1.Value);

                                }
                                else if (l == 12)
                                {
                                    float.TryParse(value1.Value.ToString(), out float x1);
                                    helper.Cells(column: i, row: j, data: x1);
                                }
                                else
                                {
                                    float.TryParse(value1.Value.ToString(), out float x1);
                                    if (x1 != 0)
                                    {
                                        i = 13;
                                        j = j + 1;
                                        float.TryParse(list34[7].ToString(), out float x2);
                                        helper.Cells(column: i, row: j, data: value1.Key);
                                        i = i + 1;
                                        helper.Cells(column: i, row: j, data: x1);
                                        i = i + 1;
                                        helper.Cells(column: i, row: j, data: x2 * x1);

                                    }

                                }

                            }
                        }

                        helper.save();

                    }
                }
            }

            catch (Exception ex) { Console.WriteLine(ex.Message); }
            {

            }
        }
        #endregion

    }
}



