//*******************************************************************
//* incbkup.js (sgrb)
//*******************************************************************
var fso = WScript.CreateObject("Scripting.FileSystemObject")
var ntobj = WScript.CreateObject("WScript.Network")
var ShellObj = WScript.CreateObject("WScript.Shell")

var objarg = WScript.Arguments
var types
var nodrv,freen,driven,namef,a0
var nndel,tedel,deltim,era
var version

version=" incbkup 1.19 (06.11.2014) "

// имя бэкап сервера по умолчанию
bkserv="\\\\priz-backup\\"

progpath=""
// узнаем есть ли RAR на компьютере либо он на сервере
whererar(progpath)

freen=0
driven=""
bkdto=""
aaz=0
deltim=""
nndel=0
tedel=""
era=""
nszera="" 
nsz=1000
dirf="dir.txt"
nlst=3
namel="priz"
pswd="priz"

//
// обработка аргументов
//
// incbkup.js [i|f|fu|?d|<число>d] [<куда писать>|-] [<число>w|<число>m|-] [e|-] [число|числоK|числоM|числоG|-] file.cfg
// Первый параметр это тип бэкапа:
//
// i        инкрементальрый по биту доступа     \
// f        полный с сбросом бита архивирования /   требуют записи на архивируемые ресурсы
//
// fu       полный без сброса бита архивирования            \
// <число>d инкрементальрый по <числу> дней предшествующих   |  не требуют записи на
//                                                           |  архивируемые ресурсы
// ?d       инкрементальрый за время с предыдущего бэкапа   /
// Значение по умолчанию "i"
//
// Вторым параметром служит устройство и каталог куда писать У:\папка,
// либо \\комп\сетеваяпапка\папка.. , если этот параметр начинается с "!", 
// то в случае отсутствия будет предринята попытка создать папку для пользователя.
// По умолчанию для бэкапа служит диск "Z" (обычно \\priz-backup\backup$), 
// диск Z: и в остальных случаях используется для определения 
// параметров на сетевой папке для бэкапа (то есть он зарезервирован).
// Если нужен третий параметр, а второй хочется пропустить, то вместо него ставят "-" (прочерк)
// Значение по умолчанию "!Z:\" 
//
// Третьим параметром задано число с последующим w или m, или y, или d,
// то указанное число означает количество недель, месяцев, лет или дней,
// которое соответствующие архивы сохраняются.
// (Если указано только число, то это означает количество дней.)
// Все остальные архивы уничтожаются, причем удаляются(сохраняются) только архивы соответствующие 
// полному бэкапу несмотря на тип бэкапа, также удаляются и соответствующие логи.
// Если нужен четвертый параметр, а третий хочется пропустить, то вместо него ставят "-" (прочерк)
// Значение по умолчанию <пусто>, то есть ничего не удалять.
//
// Следует иметь в виду что если указан параметр "е",
// то в случае полного бэкапа удаляются все ранние инкрементальные архивы и соответствующие логи
// кроме того проверяется своб.место хватит ли его на бэкап, считая р-р предыдущего полного бэкапа + 10%
// в качестве прогнозируемого размера тек.бэкапа, и в случае если места нет сокращается количество 
// сохраняемых дней, пока не станет достаточно места.
// в случае инрементального бэкапа происходит то же.
//
// Кроме если не удается определить р-р бэкапа, если указано в командной строке <число>(в Кило,Мега,Гига или просто байтах),
// то оно используется в качестве р-ра полного бэкапа
//
// Следует иметь в виду что если указан неправильный(ошибочный), то он игнорируется
// без диагностики, а качестве значения используется значение по умолчанию
// вместо этого параметра тоже может быть прочерк, 
// в любом случае последним параметром может служит имя файла-списка каталогов для бэкапа
// по умолчанию "dir.txt"

if (objarg.length>0){			// если есть параметры
  types=objarg.Item(0)			// первый параметр - тип бэкапа
  if(objarg.length>1){			// если их больше одного
    bkuppc=objarg.Item(1)		// 2-й - куда писать 
    if(bkuppc=="-")bkuppc="!Z:\\"	// если прочерк
  } else{							// или пропущен
    bkuppc="!Z:\\"					// то !Z:\\
  }
  if(objarg.length>2){			// если параметров больше 3
    deltim=objarg.Item(2)		// время сохранения бэкапов
    if(deltim=="-")deltim=""	// если прочерк
  } else {						// или пропущен
    deltim=""					// то пусто
  }
  if(objarg.length>3){
    era=objarg.Item(3)	// признак удаления
    if(era=="-"){
		era=""			// если прочерк
		nlst=4
	} else {
		if(objarg.length>4){			// время удержания сохраненых файлов
			nszera=objarg.Item(4)
			if(nszera=="-")nszera=""	// если прочерк			
		} else {
			nszera=""
		}
		nlst=5
	}	
  } else {
    era=""
	nlst=3
  }
  if(objarg.length>nlst){		// последний параметр
    dirf=objarg.Item(nlst)
  } else {
	dirf="dir.txt"  			// по умолчанию
  }
}else{
   types="i"
   bkuppc="!Z:\\"
   era=""
   deltim=""					//
}

// WScript.Echo (types+" "+bkuppc+" "+deltim)

if(types=="i"){                   //простой инкрементальный (архивный бит)
      a0=" -ao -ac "
}else if(types=="f"){             //полный с сбросом бита архивирования
      a0=" -ac "
}else if(types=="fu"){             //полный, без записи на архивируемые диски
      a0=" "
}else if(types=="?d"){            //инкрементальный, по времени предыдущего
          bkdto=types
          a0=" -tn"
}else if(/\d+d$/.test(types)){    //инкрементальный за ук.число дней
          bkdto=types
          a0=" -tn"+ bkdto
}else{
          types="i"
          a0=" -ao -ac "
}

if(deltim!=""){	//число дней удаления
    if(/\d+w$/.test(deltim)){           //за ук.число недель
            tmp=deltim.substr(0,deltim.length-1)
            nndel=parseInt(tmp)*7
            tedev="w"
    } else if (/\d+d$/.test(deltim)){   //за ук.число суток
            tmp=deltim.substr(0,deltim.length-1)
            nndel=parseInt(tmp)
            tedel="d"
    } else if (/\d+m$/.test(deltim)){   //за ук.число месяцев
            tmp=deltim.substr(0,deltim.length-1)
            nndel=parseInt(tmp)*30
            tedel="m"
    } else if (/\d+y$/.test(deltim)){   //за ук.число лет
            tmp=deltim.substr(0,deltim.length-1)
            nndel=parseInt(tmp)*365
            tedel="y"
    } else if (/\d+$/.test(deltim)){   //за ук.число суток(дней)
            nndel=parseInt(deltim)
            tedel="d"
    } else {
      nndel=0
      tedel=""
      deltim=""
    }
}

if(nndel<=0)nndel=1

if(nszera!=""){	// мин. кол-во диск. пространства
    if(/\d+G$/.test(nszera)){           //
            tmp=nszera.substr(0,nszera.length-1)
            nsz=parseInt(tmp)*1000*1000*1000
    } else if (/\d+g$/.test(nszera)){   //
            tmp=nszera.substr(0,nszera.length-1)
            nsz=parseInt(tmp)*1000*1000*1000
    } else if (/\d+m$/.test(nszera)){   //
            tmp=nszera.substr(0,nszera.length-1)
            nsz=parseInt(tmp)*1000*1000
    } else if (/\d+M$/.test(nszera)){   //
            tmp=nszera.substr(0,nszera.length-1)
            nsz=parseInt(tmp)*1000*1000
    } else if (/\d+k$/.test(nszera)){   //
            tmp=nszera.substr(0,nszera.length-1)
            nsz=parseInt(tmp)*1000
    } else if (/\d+K$/.test(nszera)){   //
            tmp=nszera.substr(0,nszera.length-1)
            nsz=parseInt(tmp)*1000
    } else if (/\d+$/.test(nszera)){   //
            nndel=parseInt(nszera)
            nsz=parseInt(tmp)*1000
    } else {
      nsz=1000
      nszera=""
    }
}  

// WScript.Echo (deltim+" "+nndel+" "+tedel)

testbk()   // проверим и подготовим место куда будем делать бэкап

//читаем настроечный файл
//tmps=pref + "\\dir.txt"
tmps=pref + "\\" + dirf

var bkupsrc=Array(20)
var nsrc
nsrc=0
var inifo
try {
  inifo = fso.opentextfile(tmps,1)
} catch(err) {
  //error
  ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp01 no folder list "  + err.number +" : "+ err.description)   //Инрформационное сообщение
  err.clear
  WScript.Quit()
}
i=0;

while(!(inifo.AtEndOfStream)){
    bkupsrc[i] = inifo.ReadLine()	// читаем файл в bkupsrc[i]
    if(bkupsrc[i].length==-1) break
    i=i+1

}
nsrc=i-1

if(types=="i"){
  outfp = "Inc"
  logss = "Incremental backup "
}else if(types=="f"){
  outfp = "Full"
  logss = "Full backup!"
}else if(types=="fu"){
  outfp = "Full"
  logss = "Full backup!"
}else{
  outfp="Inc"
  logss="Incremental for " + bkdto
}
//
TotalInputBytes = 0
// Запишем в системный Log
temp = "***"+WScript.Scriptname + " Пользователь: " +
   ntobj.UserName +
   "  Тип бэкапа: "+ logss
ShellObj.LogEvent( 4, temp)  //Инрформационное сообщение

//формируем имя лога и частично имя архива
var tempdate = new Date()
fndate = tempdate.getMonth()+1
fndate = "_" + tempdate.getDate() + "." + fndate  +
      "." + tempdate.getFullYear() + "_"
temp = tempdate.getHours()
if(temp < 10) temp = "0" + temp
fndate = fndate + temp + "."  //часы hh
temp =  tempdate.getMinutes()
if (temp < 10) temp = "0" + temp
fndate = fndate + temp        //минуты mm
logf =  pref + "\\"+ outfp +"_" +ntobj.ComputerName + "_Log" + fndate + ".txt"
elgf =  pref + "\\"+ outfp +"_" +ntobj.ComputerName + "_Err" + fndate + ".txt"

tfile=pref + "\\timestamp.txt"
//если инкр по времени то читаем timestamp (последнюю строчку)
if(types=="?d"){
  err=0
  try {
    tf=fso.opentextfile(tfile,1)
  } catch(err) {
    //error
    err=1
    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp02 timestamp error "  + err.number +" : "+ err.description)   //Инрформационное сообщение
    err.clear
    //WScript.Quit()
  }
  nday=-1
  if(err==0){
    ts=""
    while(!(tf.AtEndOfStream)){
      ts = tf.ReadLine()
      if(ts.length==-1) break
      i=i+1
    }
    tf.close()
    tts=ts.split(" ")
    // определяем сколько дней прошло с прошлого бэкапа
    olddt=Date.parse(tts[0])
    nday=Math.floor((tempdate-olddt)/86400000)
  }
  if(nday>0){
    a0=a0+nday+"d"
  }else{//если тиместампа нет то задаем очень много
    a0=a0+"10000d"
  }
}

writelog("Incbkup "+version+" Начало backup файл конфигурации: "+dirf )

//узнаем параметры устройства бэкапа
var OutDrvObj = fso.GetDrive(driven)   //!!!
OutDrvFspc = OutDrvObj.FreeSpace

switch(OutDrvObj.DriveType){
        case 0: DrvType = "Непонятное"
                  break
        case 1: DrvType = "Сьемное"
                  break
        case 2: DrvType = "Несьемное"
                  break
        case 3: DrvType = "Сетевое"
                  break
        case 4: DrvType = "CD-ROM"
                  break
        case 5: DrvType = "RAM диск"
}
writelog("  " + "Устройство бэкапа:"+ driven + " " + OutDrvObj.VolumeName +
        "(" + DrvType + ") " + OutDrvObj.FileSystem +" "+ OutDrvObj.ShareName +
        " Общий размер: " + OutDrvObj.TotalSize/1048576 + " Mб" +
        " Свободное пространство: " + OutDrvObj.FreeSpace/1048576 + " Mб")
// проверяем свободное место

if(OutDrvFspc<nsz){
    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp03 too low space on backup disk")
    writelog("Мало места на backup диске!!!") // Инрформационное сообщение
    WScript.Quit()
}

if(OutDrvFspc<2*nsz){
    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp04 warning low space on backup disk")
    writelog("Маловато места на backup диске") // Инрформационное сообщение
} 

// удаляем старые архивы если это заказано (только full!)

bksz=getprevbksz(outfp)		// узнаем сколько будем писать!?

if (bksz < -1){			
	if(outfp=="Inc"){
		bksz=getprevbksz("Full")		// все таки узнаем сколько будем писать!?
		if(bksz >0)
			bksz=0.5*bksz
		else
			bksz=nsz			// ошибка, видимо мы впервые пишем бэкап, поэтому будем безусловно писать !??
	} else {
		bksz=nsz				// если мы не можем узнать размер прошлого бэкапа, берем из командной строки, или 1000!
	}
}

if(deltim!=""){
  if(bksz>0){
	bksz=bksz*1.1				// узнаем сколько будем писать!?
	if(bksz/1048576<2000)brsz=1048576*2000	// если мало то будем считать что 2гб (спорно!!!)
	writelog("Предполагаемый р-р бэкапа: " + bksz/1048576+ " Mб")
  }  else {						// в случае ошибок открытия соответсвующих папки и файлов 
	brsz=1048576*50000	// и если нет р-ра, то это очень спорно, но пока вот так !???
	writelog("Не могу узнать р-р бэкапа, поэтому от фонаря считаем нам нужно " + bksz/1048576+ " Mб")
  }

  for(;;){
   DeleteOldFolders("Full",nndel)	// тут мы удаляем только full архивы 
   OutDrvFspc = OutDrvObj.FreeSpace
   if(era!="e"){
   	writelog("Свободно "+OutDrvFspc +" но мы чистку не заказали")
   	break
   }
   if(OutDrvFspc > bksz){
   	writelog("Места хватает "+OutDrvFspc +">" + bksz)
   	break
   }
   // если места не хватает сокращаем количество удерживаемых дней
   nndel--
   if(nndel <= 0){
	   writelog("Места фатально не хватает : "+OutDrvFspc +"<" + bksz)
	   writelog("Запись производить не будем всё кончено!!! ")
	   ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp031 too low space on backup disk")
           WScript.Quit()
	   break
   }
   // и снова на чистку
   writelog("Мало места, надо-бы ещё почистить! Сократим ещё на денёк!!")
  }

}

CopyAllFolders()
//собственно производим копирование

// удаляем старые инкрементальные архивы если делается полный и это заказано
if((outfp=="Full")&&(era=="e"))DeleteOldFolders("Inc",1)

//отмечаем дату и время окончания работы
var edate=new Date()
writelog("Работа закончена: " + edate.toLocaleString())
// отключаем z диск
if((driven=="Z")||(driven=="z")){
  try {
    ntobj.RemoveNetworkDrive("Z:",true,true)
  } catch(err) {
        //не фатальная ошибка 
        writelog("remove error Z:  " + err.number +" : "+ err.description) //Инрформационное сообщение
        err.clear
  }
}

// запишем информацию о начале и типе закончившегося бэкапа
try {
    tf=fso.opentextfile(tfile,8,true)
    fndate = tempdate.getMonth()+1
    fndate = fndate+"/"+tempdate.getDate() + "/" + tempdate.getYear()
    tf.writeline(fndate+" "+outfp)
    tf.close()
} catch(err) {
    //error
    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp05 timestamp error "  + err.number +" : "+ err.description)   //Инрформационное сообщение
    err.clear
    //WScript.Quit()
}

// shellobj.Run "%windir%\notepad " + logf
temp = "***"+ WScript.Scriptname + "cp06 бэкап завершен"+"\n"
ShellObj.LogEvent(4, temp)
//Wscript.Quit
///////////////////////////////////////// end of script //////////////////////////////////////////


// ПОДПРОГРАММЫ
//
function CopyAllFolders()
// копировать все папки в списке
{
 var indx3

 // доформировываем имя архивной папки и создаем ее
 namef=pref + "\\" +outfp+ fndate
 try {
  fso.CreateFolder(namef)
 }catch(err){
        writelog("не могу создать " + namef + "  : " + err.number +" : "+ err.description) //Инрформационное сообщение
        err.clear
 }

 // для каждого элемента списка бэкапов
 for(indx3 = 0;indx3<=nsrc; indx3++){
 // делаем его
     copy1folder(bkupsrc[indx3])
 }
 //
 var OutDrvFspcAfter
 OutDrvFspcAfter = OutDrvObj.FreeSpace
 writelog(" Свободное пространство до копирования: " + OutDrvFspc/1048576+ " Mб" +
     " После: " + OutDrvFspcAfter/1048576+ " Mб")
}

function copy1folder(dirz)
// Копировать одну папку
{
  //WScript.Echo("copy " + dirz)

  // пропускаем пустые строки и комментарии
  if(dirz=="")return(1)
  if(dirz==" ")return(1)
  if(dirz.charAt(0)=="#")return(1)
  if(dirz.charAt(0)==":"){
	l=dirz.length	// длина
	lnz=dirz.indexOf(" ")
	if(lnz==-1)return(1)
	if(l==lnz)return(1)
    namel=dirz.substr(1,lnz)				// 
	pswd=dirz.substr(lnz+1,l)				//
	return(1)	
  }

  // если строка начинается с \\ (сетевая папка)
  if(dirz.indexOf("\\\\")!=-1){
    id=1
    l=dirz.length	// длина
    //ищем 2-й "\"
    lu=dirz.indexOf("\\",3)
    if(lu!=-1){
      dll=dirz.substr(lu+1,dirz.length+1)	// имя каталога (сетевая папка)
      dln=dirz.substr(2,lu-2)				// имя компа
	  // доформировываем имя арх.папки и создаем ее FF
	  ldf=dirf.indexOf(".")-1
	  dirff=dirf.substr(1,ldf)
      ff=pref + "\\" +outfp+ fndate+"\\"+dln +"("+dirff+")"
	  try {
	   fso.CreateFolder(ff)
	  }catch(err){
	        writelog("не могу создать " + ff + "  : " + err.number +" : "+ err.description) //Инрформационное сообщение
	        err.clear
	  }     
    }else{
      dll=dirz
      dln=dirz
      writelog("не указана сетевая папка только компьютер:" + dirz ) //Инрформационное сообщение
	  return      
    }
    // 
    luu=dll.indexOf("\\")
    if(luu!=-1){

      la=luu+lu+1
    }else{
      la=l+1
      dirz=dirz+"\\"
    }
    // формируем имя архива и источник откуда писать
    if (luu!=-1){
      dln=dln+"_"+dll.substr(0,luu)+"_"+dll.substr(luu+1,dll.length-luu-1)
    }else{
      dln=dln+"_"+dll
    }

  }else{	// все остальные
    id=0
    dl  =dirz+ "\\*.*"
    dln =dirz
    ff=pref + "\\" +outfp+ fndate
  }

  if (!fso.FolderExists(dirz)){
    writelog("не могу открыть папку " + dirz ) //Инрформационное сообщение
    return
  }


  // заменим имени архива все пробелы,бэкслэши и двоеточии на подчёркивание
  arrc=""
  //WScript.Echo( dirz + "-" + dl )
  for(i=0;i<=dln.length;i++){
    c=dln.substr(i,1)
    if(c==":")c="_"
    if(c==" ")c="_"
    if(c=="\\")c="_"
    arrc+=c
  }

  //arcf = ntobj.ComputerName +"_"+arrc+  ".rar"

  arcf = arrc+  ".rar"		//	имя арх.файла
  // ff архивный каталог

  // пример - запуск командного файла:
  //    rc=shellobj.Run ("C:\WINNT\system32\cmd.exe /K file.bat", 1,True)
  // здесь он не нужен, но чтоб не забыть ...
  
  // собственно пишем бэкап папки
  var rc=-1
  ShellObj.CurrentDirectory=progpath

  wrlog="rar.exe a -ilog"+elgf+" " + a0 + " -m1 -r0 -x@" + pref + "\\exl.txt  " + ff + "\\"+arcf + " \"" + dirz + "\" "
  rc=ShellObj.Run(wrlog ,0,true)

  // результаты работы rar.exe
  
  switch(rc){
  case 0: writelog(wrlog+"rc="+rc+" : O'k")
              break
  case 1: writelog(wrlog+"rc="+rc+" : Предупредительная диагностика")
              break

  case 2: writelog(wrlog+"rc="+rc+" : Фатальная ошибка")
              break
  case 3: writelog(wrlog+"rc="+rc+" : Ошибка КС(CRC)")
              break
  case 4: writelog(wrlog+"rc="+rc+" : Заблокирован архив")
              break
  case 5: writelog(wrlog+"rc="+rc+" : Ошибка записи")
              break
  case 6: writelog(wrlog+"rc="+rc+" : Ошибка открытия")
              break
  case 7: writelog(wrlog+"rc="+rc+" : Ошибка пользователя")
              break
  case 8: writelog(wrlog+"rc="+rc+" : Ошибка памяти")
              break
  case 9: writelog(wrlog+"rc="+rc+" : Ошибка создания")
              break
  case 255: writelog(wrlog+"rc="+rc+" : Прервано пользователем")
              break
  default: writelog(wrlog+" : Неожиданный код завершения rc="+rc)
              break
  }
  return(0)
}

function DeleteOldFolders(outfpp,nddel)
// удаляeм старые архивы
// outfpp - тип бэкапа(inc/full)
// nddel  - за сколько дней не удалять
{
 var indx3
 var indx
 var tdt = new Date()
  dd=tdt.getDate()
  dd=dd-nddel
  tdt.setDate(dd)
  timdd=tdt.getTime()             // вроде как абсолютная дата
  // узнать дату усечения (абсолютную)
  //var tdd=new Date(timaa)
  writelog("удаляeм старые архивы "+outfpp+" от "+tdt.toLocaleString()+" и старше")

  var f=fso.GetFolder(pref)             // для папки бэкапов
  var ff=new Enumerator(f.SubFolders)   // взять подпапки
  var fc=new Enumerator(f.files)        // и файлы в ней
  var re=new RegExp("^"+outfpp)

  indx=0
  // для всех папок
  for(;!ff.atEnd();ff.moveNext()){
    //вытащить из имени дату
    s=ff.item().Name
    if(s.match(re)==null)continue
    n=s.indexOf("_")+1
    m=s.indexOf(".")
    da=parseInt(s.substring(n,m))       // день
    n=s.indexOf(".",m+1)
    mo=parseInt(s.substring(m+1,n))-1   // месяц
    ye=parseInt(s.substring(n+1,n+5))   // год
    var tar=new Date(ye,mo,da)  // получить дату архива(бэкапа)
    timar=tar.getTime()         // абсолютная дата архива
    // с датой больше даты усечения  пропускаем
    if(timar>timdd)continue
    // удалить содержимое:
    var fcc=fso.getFolder(pref+"\\"+s)  // архивная папка
    var fff=new Enumerator(fcc.files)   // коллекция файлов

    // для каждого элемента списка бэкапов
    for(;!fff.atEnd();fff.moveNext()){
      // делаем т.е. удаляем
      ss=fff.item().Name
      fso.DeleteFile(pref+"\\"+s+"\\"+ss,true)
      writelog(" файл "+ss+" удален")
      indx++
    }

    fso.DeleteFolder(pref+"\\"+s,true)
    writelog("*папка "+s+" удалена")
    indx++
  }

  if (indx==0){
    writelog("*архивов для удаления нет")
  } else {
    writelog("всего удалено "+indx+" файлов/папок")    
  }

  indx=0
  // для всех файлов логов с датой меньше даты усечения их удалить
  for(;!fc.atEnd();fc.moveNext()){
    // вытащить из имени дату
    s=fc.item().Name
    if(s.match(re)==null)continue
    n=s.indexOf("Log_")
    if(n == -1){
      n=s.indexOf("Err_")
      if(n == -1)continue
    }
    n+=4
    m=s.indexOf(".",n)
    da=parseInt(s.substring(n,m))       // день
    n=s.indexOf(".",m+1)
    mo=parseInt(s.substring(m+1,n))-1   // месяц
    ye=parseInt(s.substring(n+1,n+5))   // год

    var tar=new Date(ye,mo,da)  // получить дату архива(бэкапа)
    timar=tar.getTime()         // вроде как абсолютная дата
    if(timar>timdd)continue
    // с датой меньше даты усечения
    // удалить текущий (log@err) файл
    fso.DeleteFile(pref+"\\"+s,true)
    writelog("*файл "+s+" удален")
    indx++
  }

  if (indx==0){
    writelog("*файлов для удаления нет")
  } else {
        writelog("всего удалено "+indx+" файлов")    
  }

 var OutDrvFspcAfter
 OutDrvFspcAfter = OutDrvObj.FreeSpace
 writelog(" Свободное пространство после удаления: " + OutDrvFspcAfter/1048576+ " Mб")
}

function writelog(text)
// запись строки логов
{
 var logfo
 var ttr=new Date()  // получить время
//открываем лог-файл
 try {
  logfo = fso.OpenTextFile(logf,8,true)
 }catch(err){
  ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp30  cannot create log file "  + err.number +" : "+ err.description)   // Инрформационное сообщение
  err.clear
  WScript.Quit()
 }
 logfo.WriteLine(ttr.toLocaleString()+" "+text)
 logfo.close()

}

function testbk()
// проверка диска Z существования места и возможно создания каталога по имени компьютера и сл.файлов
// диск z: вообще говоря не нужен, но удобно с его помощью определять свободное пространство
{
  if (bkuppc.charAt(0)=="!"){
    createdr=true
    tmp=bkuppc.substr(1,bkuppc.length+1)
    bkuppc=tmp
  }else{
    createdr=false
  }
  //проверка и подключение умолчального диска бэкапа Z:
  if(bkuppc=="Z:\\"){
    try {
      //ntobj.MapNetworkDrive("Z:",bkserv+"backup$",true , "priz", "priz")
      ntobj.MapNetworkDrive("Z:",bkserv+"backup$",true)
    } catch(err){
      if (!(err.number==-2147024811)){//
        //не фатальная ошибка error
        ShellObj.LogEvent(4,"*"+ WScript.Scriptname + " cp1 " + err.number +" : "+ err.description) //Инрформационное сообщение
        err.clear
      }else{
        aaz=1
      }
    }
    driven="Z"
    bkuppc=bkserv+"backup$"
  }else{
    //для других вариантов указанных в командной строке
    if(bkuppc.indexOf(":")!=-1){
      //проверяем есть ли такой диск
      driven=bkuppc.charAt(0)
      nodrv=0
    }else if(bkuppc.indexOf("\\\\")!=-1){
      //это сетевое имя
      // пытаемся отсоединить диск
      try {
        ntobj.RemoveNetworkDrive("Z:",true,true)
      } catch(err) {
        //не фатальная ошибка 
        ShellObj.LogEvent(4,"*"+ WScript.Scriptname + " cp11 " +"remove error Z:  " + err.number +" : "+ err.description) //Инрформационное сообщение
        err.clear
      }

      l=bkuppc.length
      // ишем второй бэкслэш
      lu=bkuppc.indexOf("\\",3)
      if(lu!=-1){
        dll=bkuppc.substr(lu+1,bkuppc.length+1)
      }else{
        // Непонятно куда писать (не указана сетевая папка)
        ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp12 " + " not share defined")
        err.clear
        WScript.Quit()
      }
      luu=dll.indexOf("\\")
      if(luu!=-1){
        la=luu+lu+1
      } else{
        la=l
      }

      // пытаемся подключить папку к диску z: хотя это необязательно
      try {
        ntobj.MapNetworkDrive("Z:",bkuppc.substr(0,la),true)
      } catch(err){
        if (!(err.number==-2147024811)){//
          //не фатальная ошибка error
          ShellObj.LogEvent(4,"*"+ WScript.Scriptname + " cp13 " + err.number +" : "+ err.description) //Инрформационное сообщение
          err.clear
        }else{
          aaz=1
        }
      }
      driven="z"
      nodrv=0
    }else{
      //по умолчанию считаем что это D
      driven="D"
      nodrv=1
    }
  }

  if(bkuppc.charAt(bkuppc.length-1)!="\\")pre="\\"
  else pre=""

  if(nodrv==1) drv=driven +":\\"
  else  drv=""

  //построим префикс каталога где будут храниться бэкапы,логи и т.п.
  pref=drv+bkuppc+pre +ntobj.ComputerName

  if (fso.FolderExists(pref)) {
    //MsgBox "Папка существует"
    if(!fso.FileExists(pref+"\\"+dirf)){
      ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp17 " + "cannot file "+ pref+"\\"+dirf +" open")
      WScript.Quit()
    }
  }else if (createdr){
      //создать папку и файлы
      createus(drv+bkuppc+pre,ntobj.ComputerName)
  }else{
      // Непонятно куда писать(отсутствует папка)
      ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp18 " + "cannot folder open")
      WScript.Quit()
  }

  // после всех манипуляций, если нужен Z, находится ли он среди дисков системы
  if(driven=="Z"){
    nozfind()
  }

}

function createus(pr,ne)
// создать каталог пользователя с именем его компльютера и прототипы файлов dir.txt и exl.txt
{
  try {
    f = fso.GetFolder(pr)
    fc = f.SubFolders
    fc.Add (ne)
  }catch(err){
    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp20 create folder" + err.number +" : "+ err.description)
    err.clear
    WScript.Quit()
  }


  try {
    ts = fso.OpenTextFile(pref+"\\dir.txt", 2, true)
    // проверка в какой системе работаем 
	if (fso.FolderExists("c:\\Users\\")) {
		//Папка существует - windows7/vista/8
		ts.WriteLine ("c:\\Users\\"+ntobj.UserName)
	}else{
		//папки нет windows XP
		ts.WriteLine ("c:\\Documents And Settings\\"+ntobj.UserName)
	}
    ts.WriteLine( "#d:\\" )
    ts.Close()
  }catch(err){
    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp21 create dir.txt" + err.number +" : "+ err.description)
    err.clear
    WScript.Quit()
  }

  try {
    ts = fso.OpenTextFile(pref+"\\exl.txt", 2, true)
    ts.WriteLine( "*.bak" )
    ts.WriteLine ("*.jpg")
    ts.Close()
  }catch(err){

    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp22 create exl.txt" + err.number +" : "+ err.description)
    err.clear
    WScript.Quit()
  }
}

function nozfind()
//поиск среди сетевых дисков етого самого диска Z: зачем? да на всякий случай чтоб не забыть как это делать
{
  var colDrives = ntobj.EnumNetworkDrives();
  var strMsg=""

  try {
    if (colDrives.length != 0) {
      for (i = 0; i < colDrives.length; i += 2) {
        strMsg = strMsg + "\n" + colDrives(i) + "\t" + colDrives(i + 1);
        if(colDrives(i)=="Z:"){
          driven="Z"
          freen=fso.GetDrive(driven).FreeSpace/1048576
          break
        }
      }
    }
  } catch(err){
    //не фатальная ошибка
    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp41 "+ err.number +" : "+ err.description) //Инрформационное сообщение
    err.clear
  }
  if(driven!="Z"){
    // нет Z
    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp42  no Z: drive")                         //Инрформационное сообщение
    WScript.Quit()
  }

}

function whererar()
// где же RAR
{
  path="C:\\Program Files\\WinRAR"

  if (fso.FolderExists(path)) {
	// на текущем компьютере
    //MsgBox "Папка существует"
    if(fso.FileExists(path+"\\rar.exe")){
      progpath=path
      return		

    }
  }else {
	// на backup-сервере
    path=bkserv+"backbin"
    if (fso.FileExists(path+"\\rar.exe")){
      progpath=bkserv+"backbin"
      return
    }else{
        // не могу найти rar
        ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp51 " + "cannot find RAR")
        WScript.Quit()
    }
  }

}

function getprevbksz(outfpp)
// получить размер предыдущего бэкапа
// outfpp - тип (Full/Inc) бэкапа
{
  
  tt=""
  err=0
  // попробуем найти последний бэкап
  try {
    tf=fso.opentextfile(tfile,1)
  } catch(err) {
    //error
    err=1
    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp021 timestamp error "  + err.number +" : "+ err.description)   //Инрформационное сообщение
    err.clear
    //WScript.Quit()
  }
  
  if(err==0){
	 // в файле протоколирования времен запуска ищем последний бэкап
    ts=""
    i=0
    while(!(tf.AtEndOfStream)){
      ts = tf.ReadLine()
      if(ts.length==-1) break
      tts=ts.split(" ")
      if(tts[1]==outfpp){
		  tt=tts
	  }
      i=i+1
    }
    tf.close()
    if(tt==""){
		lastfull=""							//не нашли
		return -2
	}
    var f1=fso.GetFolder(pref)             	// для папки бэкапов
    var fo=new Enumerator(f1.SubFolders)   	// взять подпапки
    // 1-day,0-month,2-year
    tts=tt[0].split("/")

    //имя архива
    fndat = "_" + tts[1] + "." + tts[0]  + "." + tts[2] + "_"
    nameff=tt[1]+ fndat
    var re=new RegExp("^"+nameff)

    indx=0
    i1=0
    // для всех подпапок ищем нашу
    for(;!fo.atEnd();fo.moveNext()){
      //вытащить имя
      s=fo.item().Name
      if(s.match(re)!=null){
		  // если нашли!!!
		  i1=1	
		  break
      }
      indx++
    }
    if(i1==0){ // не нашли
      lastfull=""
	   return -3
	}
    //какой же её размер
	lastfull=pref+"\\"+s    
    var f = fso.GetFolder(lastfull);
    var x = f.Size;	// а вот такой
    return x
  }
  lastfull=""
  return -1
}

