//Для работы скрипта понадобится заранее создать лист с названием "Управление" и заменить ссылку с адресом портала на свою 
 
 var ss = SpreadsheetApp.getActiveSpreadsheet();      // возвращаем текущую активную таблицу 
 // удаление всех листов, кроме первого перед заполнением
 var b = 0;
 while(ss.getSheets()[b]!=null) b++;
 while (b!=1){
   b--;
   ss.deleteSheet(ss.getSheets()[1]);
 }

  var response = UrlFetchApp.fetch("*ваша ссылка на Битрикс*"); //получаем доступ к данным по адресу портала
  var data  = JSON.parse(response); // преобразованный файл с данными
  var n = data.result.tasks.length; // записываем количество задач (длина массива)

 var row_count = [];   // массив для учета кол-ва записей каждого листа 
  var k =[];              // массив отступов, на сколько нужно сдвигать строки на каждом листе (при разделении задач по месяцам)
  var cm = [];           // массив с текущим месяцем каждого листа
  var group = [];       // массив групп
  
 

  // Таблица для управления
  var sheet = ss.getSheetByName("Управление"); //подключаемся по имени листа таблицы к нужному
  var rang = sheet.getRange("A1:I1");  
rang.setFontFamily("Times New Roman")
     .setFontSize(14)
     .setBackground("#87CEFA")
     .setHorizontalAlignment("center")
     .setFontWeight("bold");
rang = sheet.getRange("A2:D1000");
rang.setFontFamily("Times New Roman")
     .setFontSize(14);
sheet.getRange("A2:B1000")         
     .setHorizontalAlignment("left");
sheet.getRange("C2:I1000")         
     .setHorizontalAlignment("right");     
for (var i = 1; i <= 4; i++) { 
  var column_title = sheet.getRange(1,i);
switch (i) {
  case 1:
  column_title.setValue("Сотрудник");
  break;
  case 2:
  column_title.setValue("Месяц");
  break;
  case 3:
  column_title.setValue("Коэффициент");
  break;
  case 4:
  column_title.setValue("Ставка");
  break;
  }
 sheet.autoResizeColumn(i);
 }
  sheet.getRange(1,9).setValue("С какой даты выводить задачи (дд.мм.гггг)");

 //Кнопка обновления полей
 var button = ss.getActiveSheet().getDrawings(); 
 button.find(drawing => drawing.setPosition(8,6,0,0)) // положение - левый верхний угол 8 строка, 6 столбец

  ss.insertSheet("Все проекты за период");
  SpreadsheetApp.flush();
 var pr = ss.getSheetByName("Все проекты за период");
SpreadsheetApp.flush();   //создаем лист для всех проектов
  ss.insertSheet("Задачи без проекта"); //создаем лист для задач без проекта

  k.push(0);
  cm.push(0);
  row_count.push(0);

  // создание листов с проектами
  for (var i = 0; i < n; i++) {
    // если имя группы не пустое и листа с таким названием нет
    
   
SpreadsheetApp.flush();
    if (data.result.tasks[i].group.name != null && ss.getSheetByName(data.result.tasks[i].group.name)== null){
   SpreadsheetApp.flush();

    sheet_name = data.result.tasks[i].group.name;
    sheet = ss.insertSheet(sheet_name);   // создаем лист с названием проекта
  }
}
  // подсчет количества листов
  b=0;
  while(ss.getSheets()[b]!=null) b++;

  /* поиск максимального id проекта
  var max_id = 0;
  for (var i = 0; i < n; i++) {  
  var sr = data.result.tasks[i].group.id;
  if (sr > max_id) {
    max_id =sr;
  }
}
// количество проектов
 var q_proj = max_id/2; */  

// добавление (элементов) в массив по количеству листов-1
for (var i = 0; i < b-1; i++) {
     row_count.push(0); // добавление элемента в массив количества записей 
     k.push(0);         // добавление элемента в массив отступов от строк
     cm.push(0);        // добавление элемента в массив текущих месяцев      
  }

var period = ss.getSheetByName("Управление").getRange(2,9).getValue();

function Transfer(){  
  for (var i = 0; i < n; i++) {
    var d = data.result.tasks[i].createdDate;
   if (Utilities.formatDate(new Date(d), "GMT+08", "DD.MM.YYYY")>= Utilities.formatDate(new Date(period), "GMT+08", "DD.MM.YYYY")){
  sheet = ss.getSheetByName(data.result.tasks[i].group.name);    // выбор таблицы с с именем проекта (проход по списку всех задач)
  if (sheet == null) sheet=ss.getSheetByName("Задачи без проекта"); // если такого нет, то выбираем лист для задач без проекта (Задачи без проекта)
  ss.setActiveSheet(sheet); // переход на лист для записи задач

/* 
"ss.getActiveSheet().getIndex()-1" - индекс активного листа -1 (-1 для получения номера элемента в массиве)
row_count[ss.getActiveSheet().getIndex()-1] - количество записей на активном листе
k[ss.getActiveSheet().getIndex()-1] - количество строк отступа для активного листа
*/
  var index = row_count[ss.getActiveSheet().getIndex()-1]-k[ss.getActiveSheet().getIndex()-1];    // количество задач на листе (с учетом отступа при добавлении месяца)
  var kk = ss.getActiveSheet().getIndex()-1; // присваиваем для использования как номера элемента в массиве, k[kk] - сдвиг на активном листе (из-за месяца)
  
 
  Created_date(i,index,k[kk]);
  Closed_date(i,index,k[kk]);
  Title(i,index,k[kk]);
  Responsible(i,index,k[kk]);
  Time_estimate(i,index,k[kk]);
  Time_spent_in_logs(i,index,k[kk]);
  Effect(index+k[kk],ss.getActiveSheet());
  Effect(index+k[kk],ss.getSheetByName("Все проекты за период"));
 // Skill(index+k[kk]);
  row_count[ss.getActiveSheet().getIndex()-1]++;   // увеличение кол-ва записей в активном листе
  pr.getRange(row_count[0]+2,11).setValue(String(ss.getActiveSheet().getSheetName()));
  row_count[0]++; 
  }
  }
  for (var i = 1; i <  row_count.length; i++) {   // для каждого листа
 sheet=ss.getSheets()[i];                        //переходим на лист (i)   
 Skill(i);                               
 Text_format(i);
 Total_effect(i);
 Table_header();
 Months(i);
 // Date_format(i);
 }
}

// параметр даты для вывода только месяца словом
var options = {           
  month: 'long'
};

function Created_date(i,index,kk) { 
  var row = sheet.getRange(index +2+kk,1); 
  var created_data = data.result.tasks[i].createdDate;
  if (new Date(created_data).getMonth()+1>cm[ss.getActiveSheet().getIndex()-1]){
    cm[ss.getActiveSheet().getIndex()-1] = new Date(created_data).getMonth()+1; 
  //if (Utilities.formatDate(new Date(created_data), "GMT+08", "MM") < Utilities.formatDate(new Date(), "GMT+08", "MM")){
    k[ss.getActiveSheet().getIndex()-1]++; 
    row_count[ss.getActiveSheet().getIndex()-1]++; 
    row.setValue(new Date(created_data).toLocaleString("ru",options)).setFontWeight("bold");
    row = sheet.getRange(index +3+kk,1);
  }
  created_data = Utilities.formatDate(new Date(created_data), "GMT+08", "dd.MM.YYYY");
 row.setValue(String(created_data));
 pr.getRange(row_count[0]+2,1).setValue(String(created_data));
  }

function Closed_date(i,index,kk) {
  var row = sheet.getRange(index+2+kk,2); 
  var closed_data = data.result.tasks[i].closedDate;
  if (closed_data == null) {
    closed_data = "Не завершена";
    row.setValue(String(closed_data));
    pr.getRange(row_count[0]+2,2).setValue(String(closed_data));
    } 
  else {
  closed_data =  Utilities.formatDate(new Date(closed_data), "GMT+08", "dd.MM.YYYY");
  row.setValue(closed_data);
  pr.getRange(row_count[0]+2,2).setValue(closed_data); 
  }
}

function Title(i,index,kk) {
  var row = sheet.getRange(index+2+kk,3); 
  var titles = data.result.tasks[i].title;
  row.setValue(titles); 
  pr.getRange(row_count[0]+2,3).setValue(titles); 
  }

function Responsible(i,index,kk) {
  var row = sheet.getRange(index+2+kk,4); 
  var resp = data.result.tasks[i].responsible.name;
  row.setValue(resp); 
  pr.getRange(row_count[0]+2,4).setValue(resp);
  }

function Time_estimate(i,index,kk) {
  var row = sheet.getRange(index+2+kk,5); 
  var time_plan = data.result.tasks[i].timeEstimate;
  row.setValue(Math.round(time_plan/3600)); 
  pr.getRange(row_count[0]+2,5).setValue(Math.round(time_plan/3600));
  }

function Time_spent_in_logs(i,index,kk) {
  var row = sheet.getRange(index+2+kk,6); 
  var time_fact = data.result.tasks[i].timeSpentInLogs;
  row.setValue(Math.ceil(time_fact/3600));
  pr.getRange(row_count[0]+2,6).setValue(Math.ceil(time_fact/3600)); 
  }

function Effect(index,sht) {
x=index;
  if (sht.getSheetName() == "Все проекты за период")
 x=row_count[0];
var timer = sht.getRange(x+2,5).getValue();
var fact = sht.getRange(x+2,6).getValue();
var row = sht.getRange(x+2,7); 
if (sht.getRange(x+2,2).getValue()!="Не завершена"){
 row.setValue(Math.round(Number(timer)/Number(fact)*100));
 if  (Number(fact) == 0)   
   row.setBackground("white");
  else if (Number(fact) > Number(timer))
 row.setBackground("#DC143C"); // красный
  else if (Number(fact) == Number(timer))
 row.setBackground("#FFD700"); // желтый
   else row.setBackground("#00FA9A");  //зеленый
}
  }

function Skill(index)  { 
var table = ss.getSheetByName("Управление");
var x=row_count[index];
if (sheet.getSheetName()== "Все проекты за период")
 x=row_count[0];
  for (var i = 1; i <= x+1; i++) {
var worker = sheet.getRange(i+1,4).getValue();
var month = new Date(sheet.getRange(i+1,1).getValue()).toLocaleString("ru",options);
var level = sheet.getRange(i+1,8);
var price1 = sheet.getRange(i+1,9);
var price2 = sheet.getRange(i+1,10);
var j = 1;
while (table.getRange(j+1,1).getValue() != "") {
if (worker == table.getRange(j+1,1).getValue() && month == table.getRange(j+1,2).getValue()){
level.setValue(String(table.getRange(j+1,3).getValue()));
price1.setValue(String(Number(table.getRange(j+1,3).getValue())*Number(table.getRange(j+1,4).getValue())*Number(sheet.getRange(i+1,5).getValue())));
price2.setValue(String(Number(table.getRange(j+1,3).getValue())*Number(table.getRange(j+1,4).getValue())*Number(sheet.getRange(i+1,6).getValue())));;
}
j++;
 }
  }
} 

function Table_header() {
var range = sheet.getRange("A1:J1");  
range.setFontFamily("Times New Roman")
     .setFontSize(14)
     .setBackground("#87CEFA")
     .setHorizontalAlignment("center")
     .setFontWeight("bold");

for (var i = 1; i <= 10; i++) { 
  var column_title = sheet.getRange(1,i);
switch (i) {
  case 1:
  column_title.setValue("Дата поступления");
  break;
  case 2:
  column_title.setValue("Дата выполнения");
  break;
  case 3:
  column_title.setValue("Работы");
  break;
  case 4:
  column_title.setValue("Специалист");
  break;
  case 5:
  column_title.setValue("Оценка,часы");
  break;
  case 6:
  column_title.setValue("Факт,часы");
  break;
  case 7:
  column_title.setValue("Эффективность, %");
  break;
  case 8:
  column_title.setValue("Коэффициент");
  break;
  case 9:
  column_title.setValue("Оплата, оценка");
  break;
  case 10:
  column_title.setValue("Оплата, факт");
  break;
   }
sheet.autoResizeColumn(i);
 }
 if (sheet.getSheetName()=="Все проекты за период")
  sheet.getRange(1,11)
     .setValue("Проект")
     .setFontFamily("Times New Roman")
     .setFontSize(14)
     .setBackground("#87CEFA")
     .setHorizontalAlignment("center")
     .setFontWeight("bold")
  .setBackground("#87CEFA");
  sheet.autoResizeColumn(11);
}

function Text_format(index)  {
var s = sheet.getRange("A2:K1000");                    // полный диапазон заполняемых полей
x = row_count[index];
if (sheet.getSheetName()== "Все проекты за период")
 x=n;
s = s.setFontFamily("Times New Roman");
s = s.setFontSize(14);
s = sheet.getRange("B2:K1000").setFontWeight("normal");
   for (var i = 1; i <= x+1; i++) {                   // для кол-ва задач +1
     for (var k = 1; k <= 4; k++){                     //берём область из первых 4 столбцов
  var text = sheet.getRange(i+1,k);                   
  text = text.setHorizontalAlignment("left");         //выравнивание по левому краю
   }
     for (var k = 5; k <= 10; k++){                    //берём область с 5 по 10 столбец
  var hours = sheet.getRange(i+1,k);                 
  hours = hours.setHorizontalAlignment("right");      //выравнивание по правому краю
   } 
   sheet.getRange(i+1,7)
   .setFontColor("white")
   .setHorizontalAlignment("center");                                                                                                       
   for (var k = 11; k <= 11; k++){                     //берём 11 столбец
  sheet.getRange(i+1,k);                   
  text.setHorizontalAlignment("left");         //выравнивание по левому краю
   }
  }
}  
function Months(index) {
  var c = 0;
  var row = "";
  var first = 0;
  for (var i = 1; i <= row_count[index]+1; i++) {
 row = sheet.getRange(i,2).getValue(); 
  if (row=="") {
    c++;
    if (first==0) first=i;
  }
}
  var f = first+1;
  for (var i = 0; i < c; i++){
  for (var j = f+1; j <= row_count[index]+2; j++) {
    row = sheet.getRange(j,2).getValue();
    if (row == "") {    
  var range = sheet.getRange("A"+String(f)+":K"+String(j-1));
 range.shiftRowGroupDepth(1);
group.push(sheet.getRowGroup(i+2, 1));
if (i!=c-1) 
group[group.length-1].collapse();
f=j+1;
break;
    }
   }
  }
}

function Total_effect (index) {
 var plan_total = 0;
 var fact_total = 0;
 var x = row_count[index];
 if (sheet.getSheetName()== "Все проекты за период")
 x=row_count[0];
  for (var i = 2; i <= x+1; i++) {
   if (sheet.getRange(i,2).getValue()!="Не завершена"){
if(sheet.getRange(i,5).getValue()!="")
 plan_total = plan_total+ sheet.getRange(i,5).getValue();
if(sheet.getRange(i,6).getValue()!="") 
fact_total = fact_total+ sheet.getRange(i,6).getValue();
if (fact_total!=0) sheet.getRange(x+4,7)
.setValue(Math.round(Number(plan_total)/Number(fact_total)*100))
.setFontWeight("bold")
.setFontColor("white")
.setHorizontalAlignment("center");
 }
  }   
sheet.getRange(x+4,4).setValue("Общая эффективность").setFontWeight("bold");  // печать на 2 строчки ниже таблицы
sheet.getRange(x+4,5).setValue(String(plan_total)).setFontWeight("bold");   
sheet.getRange(x+4,6).setValue(String(fact_total)).setFontWeight("bold");
if (fact_total == 0)   
   sheet.getRange(x+4,7).setBackground("white");
  else if (fact_total > plan_total)
 sheet.getRange(x+4,7).setBackground("#DC143C"); // красный
  else if (fact_total == plan_total)
 sheet.getRange(x+4,7).setBackground("#FFD700"); // желтый
   else sheet.getRange(x+4,7).setBackground("#00FA9A");  //зеленый
}
