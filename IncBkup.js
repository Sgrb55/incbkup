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

// ��� ����� ������� �� ���������
bkserv="\\\\priz-backup\\"

progpath=""
// ������ ���� �� RAR �� ���������� ���� �� �� �������
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
// ��������� ����������
//
// incbkup.js [i|f|fu|?d|<�����>d] [<���� ������>|-] [<�����>w|<�����>m|-] [e|-] [�����|�����K|�����M|�����G|-] file.cfg
// ������ �������� ��� ��� ������:
//
// i        ��������������� �� ���� �������     \
// f        ������ � ������� ���� ������������� /   ������� ������ �� ������������ �������
//
// fu       ������ ��� ������ ���� �������������            \
// <�����>d ��������������� �� <�����> ���� ��������������   |  �� ������� ������ ��
//                                                           |  ������������ �������
// ?d       ��������������� �� ����� � ����������� ������   /
// �������� �� ��������� "i"
//
// ������ ���������� ������ ���������� � ������� ���� ������ �:\�����,
// ���� \\����\������������\�����.. , ���� ���� �������� ���������� � "!", 
// �� � ������ ���������� ����� ���������� ������� ������� ����� ��� ������������.
// �� ��������� ��� ������ ������ ���� "Z" (������ \\priz-backup\backup$), 
// ���� Z: � � ��������� ������� ������������ ��� ����������� 
// ���������� �� ������� ����� ��� ������ (�� ���� �� ��������������).
// ���� ����� ������ ��������, � ������ ������� ����������, �� ������ ���� ������ "-" (�������)
// �������� �� ��������� "!Z:\" 
//
// ������� ���������� ������ ����� � ����������� w ��� m, ��� y, ��� d,
// �� ��������� ����� �������� ���������� ������, �������, ��� ��� ����,
// ������� ��������������� ������ �����������.
// (���� ������� ������ �����, �� ��� �������� ���������� ����.)
// ��� ��������� ������ ������������, ������ ���������(�����������) ������ ������ ��������������� 
// ������� ������ �������� �� ��� ������, ����� ��������� � ��������������� ����.
// ���� ����� ��������� ��������, � ������ ������� ����������, �� ������ ���� ������ "-" (�������)
// �������� �� ��������� <�����>, �� ���� ������ �� �������.
//
// ������� ����� � ���� ��� ���� ������ �������� "�",
// �� � ������ ������� ������ ��������� ��� ������ ��������������� ������ � ��������������� ����
// ����� ���� ����������� ����.����� ������ �� ��� �� �����, ������ �-� ����������� ������� ������ + 10%
// � �������� ��������������� ������� ���.������, � � ������ ���� ����� ��� ����������� ���������� 
// ����������� ����, ���� �� ������ ���������� �����.
// � ������ ��������������� ������ ���������� �� ��.
//
// ����� ���� �� ������� ���������� �-� ������, ���� ������� � ��������� ������ <�����>(� ����,����,���� ��� ������ ������),
// �� ��� ������������ � �������� �-�� ������� ������
//
// ������� ����� � ���� ��� ���� ������ ������������(���������), �� �� ������������
// ��� �����������, � �������� �������� ������������ �������� �� ���������
// ������ ����� ��������� ���� ����� ���� �������, 
// � ����� ������ ��������� ���������� ����� ������ ��� �����-������ ��������� ��� ������
// �� ��������� "dir.txt"

if (objarg.length>0){			// ���� ���� ���������
  types=objarg.Item(0)			// ������ �������� - ��� ������
  if(objarg.length>1){			// ���� �� ������ ������
    bkuppc=objarg.Item(1)		// 2-� - ���� ������ 
    if(bkuppc=="-")bkuppc="!Z:\\"	// ���� �������
  } else{							// ��� ��������
    bkuppc="!Z:\\"					// �� !Z:\\
  }
  if(objarg.length>2){			// ���� ���������� ������ 3
    deltim=objarg.Item(2)		// ����� ���������� �������
    if(deltim=="-")deltim=""	// ���� �������
  } else {						// ��� ��������
    deltim=""					// �� �����
  }
  if(objarg.length>3){
    era=objarg.Item(3)	// ������� ��������
    if(era=="-"){
		era=""			// ���� �������
		nlst=4
	} else {
		if(objarg.length>4){			// ����� ��������� ���������� ������
			nszera=objarg.Item(4)
			if(nszera=="-")nszera=""	// ���� �������			
		} else {
			nszera=""
		}
		nlst=5
	}	
  } else {
    era=""
	nlst=3
  }
  if(objarg.length>nlst){		// ��������� ��������
    dirf=objarg.Item(nlst)
  } else {
	dirf="dir.txt"  			// �� ���������
  }
}else{
   types="i"
   bkuppc="!Z:\\"
   era=""
   deltim=""					//
}

// WScript.Echo (types+" "+bkuppc+" "+deltim)

if(types=="i"){                   //������� ��������������� (�������� ���)
      a0=" -ao -ac "
}else if(types=="f"){             //������ � ������� ���� �������������
      a0=" -ac "
}else if(types=="fu"){             //������, ��� ������ �� ������������ �����
      a0=" "
}else if(types=="?d"){            //���������������, �� ������� �����������
          bkdto=types
          a0=" -tn"
}else if(/\d+d$/.test(types)){    //��������������� �� ��.����� ����
          bkdto=types
          a0=" -tn"+ bkdto
}else{
          types="i"
          a0=" -ao -ac "
}

if(deltim!=""){	//����� ���� ��������
    if(/\d+w$/.test(deltim)){           //�� ��.����� ������
            tmp=deltim.substr(0,deltim.length-1)
            nndel=parseInt(tmp)*7
            tedev="w"
    } else if (/\d+d$/.test(deltim)){   //�� ��.����� �����
            tmp=deltim.substr(0,deltim.length-1)
            nndel=parseInt(tmp)
            tedel="d"
    } else if (/\d+m$/.test(deltim)){   //�� ��.����� �������
            tmp=deltim.substr(0,deltim.length-1)
            nndel=parseInt(tmp)*30
            tedel="m"
    } else if (/\d+y$/.test(deltim)){   //�� ��.����� ���
            tmp=deltim.substr(0,deltim.length-1)
            nndel=parseInt(tmp)*365
            tedel="y"
    } else if (/\d+$/.test(deltim)){   //�� ��.����� �����(����)
            nndel=parseInt(deltim)
            tedel="d"
    } else {
      nndel=0
      tedel=""
      deltim=""
    }
}

if(nndel<=0)nndel=1

if(nszera!=""){	// ���. ���-�� ����. ������������
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

testbk()   // �������� � ���������� ����� ���� ����� ������ �����

//������ ����������� ����
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
  ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp01 no folder list "  + err.number +" : "+ err.description)   //��������������� ���������
  err.clear
  WScript.Quit()
}
i=0;

while(!(inifo.AtEndOfStream)){
    bkupsrc[i] = inifo.ReadLine()	// ������ ���� � bkupsrc[i]
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
// ������� � ��������� Log
temp = "***"+WScript.Scriptname + " ������������: " +
   ntobj.UserName +
   "  ��� ������: "+ logss
ShellObj.LogEvent( 4, temp)  //��������������� ���������

//��������� ��� ���� � �������� ��� ������
var tempdate = new Date()
fndate = tempdate.getMonth()+1
fndate = "_" + tempdate.getDate() + "." + fndate  +
      "." + tempdate.getFullYear() + "_"
temp = tempdate.getHours()
if(temp < 10) temp = "0" + temp
fndate = fndate + temp + "."  //���� hh
temp =  tempdate.getMinutes()
if (temp < 10) temp = "0" + temp
fndate = fndate + temp        //������ mm
logf =  pref + "\\"+ outfp +"_" +ntobj.ComputerName + "_Log" + fndate + ".txt"
elgf =  pref + "\\"+ outfp +"_" +ntobj.ComputerName + "_Err" + fndate + ".txt"

tfile=pref + "\\timestamp.txt"
//���� ���� �� ������� �� ������ timestamp (��������� �������)
if(types=="?d"){
  err=0
  try {
    tf=fso.opentextfile(tfile,1)
  } catch(err) {
    //error
    err=1
    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp02 timestamp error "  + err.number +" : "+ err.description)   //��������������� ���������
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
    // ���������� ������� ���� ������ � �������� ������
    olddt=Date.parse(tts[0])
    nday=Math.floor((tempdate-olddt)/86400000)
  }
  if(nday>0){
    a0=a0+nday+"d"
  }else{//���� ���������� ��� �� ������ ����� �����
    a0=a0+"10000d"
  }
}

writelog("Incbkup "+version+" ������ backup ���� ������������: "+dirf )

//������ ��������� ���������� ������
var OutDrvObj = fso.GetDrive(driven)   //!!!
OutDrvFspc = OutDrvObj.FreeSpace

switch(OutDrvObj.DriveType){
        case 0: DrvType = "����������"
                  break
        case 1: DrvType = "�������"
                  break
        case 2: DrvType = "���������"
                  break
        case 3: DrvType = "�������"
                  break
        case 4: DrvType = "CD-ROM"
                  break
        case 5: DrvType = "RAM ����"
}
writelog("  " + "���������� ������:"+ driven + " " + OutDrvObj.VolumeName +
        "(" + DrvType + ") " + OutDrvObj.FileSystem +" "+ OutDrvObj.ShareName +
        " ����� ������: " + OutDrvObj.TotalSize/1048576 + " M�" +
        " ��������� ������������: " + OutDrvObj.FreeSpace/1048576 + " M�")
// ��������� ��������� �����

if(OutDrvFspc<nsz){
    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp03 too low space on backup disk")
    writelog("���� ����� �� backup �����!!!") // ��������������� ���������
    WScript.Quit()
}

if(OutDrvFspc<2*nsz){
    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp04 warning low space on backup disk")
    writelog("�������� ����� �� backup �����") // ��������������� ���������
} 

// ������� ������ ������ ���� ��� �������� (������ full!)

bksz=getprevbksz(outfp)		// ������ ������� ����� ������!?

if (bksz < -1){			
	if(outfp=="Inc"){
		bksz=getprevbksz("Full")		// ��� ���� ������ ������� ����� ������!?
		if(bksz >0)
			bksz=0.5*bksz
		else
			bksz=nsz			// ������, ������ �� ������� ����� �����, ������� ����� ���������� ������ !??
	} else {
		bksz=nsz				// ���� �� �� ����� ������ ������ �������� ������, ����� �� ��������� ������, ��� 1000!
	}
}

if(deltim!=""){
  if(bksz>0){
	bksz=bksz*1.1				// ������ ������� ����� ������!?
	if(bksz/1048576<2000)brsz=1048576*2000	// ���� ���� �� ����� ������� ��� 2�� (������!!!)
	writelog("�������������� �-� ������: " + bksz/1048576+ " M�")
  }  else {						// � ������ ������ �������� �������������� ����� � ������ 
	brsz=1048576*50000	// � ���� ��� �-��, �� ��� ����� ������, �� ���� ��� ��� !???
	writelog("�� ���� ������ �-� ������, ������� �� ������ ������� ��� ����� " + bksz/1048576+ " M�")
  }

  for(;;){
   DeleteOldFolders("Full",nndel)	// ��� �� ������� ������ full ������ 
   OutDrvFspc = OutDrvObj.FreeSpace
   if(era!="e"){
   	writelog("�������� "+OutDrvFspc +" �� �� ������ �� ��������")
   	break
   }
   if(OutDrvFspc > bksz){
   	writelog("����� ������� "+OutDrvFspc +">" + bksz)
   	break
   }
   // ���� ����� �� ������� ��������� ���������� ������������ ����
   nndel--
   if(nndel <= 0){
	   writelog("����� �������� �� ������� : "+OutDrvFspc +"<" + bksz)
	   writelog("������ ����������� �� ����� �� �������!!! ")
	   ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp031 too low space on backup disk")
           WScript.Quit()
	   break
   }
   // � ����� �� ������
   writelog("���� �����, ����-�� ��� ���������! �������� ��� �� ����!!")
  }

}

CopyAllFolders()
//���������� ���������� �����������

// ������� ������ ��������������� ������ ���� �������� ������ � ��� ��������
if((outfp=="Full")&&(era=="e"))DeleteOldFolders("Inc",1)

//�������� ���� � ����� ��������� ������
var edate=new Date()
writelog("������ ���������: " + edate.toLocaleString())
// ��������� z ����
if((driven=="Z")||(driven=="z")){
  try {
    ntobj.RemoveNetworkDrive("Z:",true,true)
  } catch(err) {
        //�� ��������� ������ 
        writelog("remove error Z:  " + err.number +" : "+ err.description) //��������������� ���������
        err.clear
  }
}

// ������� ���������� � ������ � ���� �������������� ������
try {
    tf=fso.opentextfile(tfile,8,true)
    fndate = tempdate.getMonth()+1
    fndate = fndate+"/"+tempdate.getDate() + "/" + tempdate.getYear()
    tf.writeline(fndate+" "+outfp)
    tf.close()
} catch(err) {
    //error
    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp05 timestamp error "  + err.number +" : "+ err.description)   //��������������� ���������
    err.clear
    //WScript.Quit()
}

// shellobj.Run "%windir%\notepad " + logf
temp = "***"+ WScript.Scriptname + "cp06 ����� ��������"+"\n"
ShellObj.LogEvent(4, temp)
//Wscript.Quit
///////////////////////////////////////// end of script //////////////////////////////////////////


// ������������
//
function CopyAllFolders()
// ���������� ��� ����� � ������
{
 var indx3

 // ��������������� ��� �������� ����� � ������� ��
 namef=pref + "\\" +outfp+ fndate
 try {
  fso.CreateFolder(namef)
 }catch(err){
        writelog("�� ���� ������� " + namef + "  : " + err.number +" : "+ err.description) //��������������� ���������
        err.clear
 }

 // ��� ������� �������� ������ �������
 for(indx3 = 0;indx3<=nsrc; indx3++){
 // ������ ���
     copy1folder(bkupsrc[indx3])
 }
 //
 var OutDrvFspcAfter
 OutDrvFspcAfter = OutDrvObj.FreeSpace
 writelog(" ��������� ������������ �� �����������: " + OutDrvFspc/1048576+ " M�" +
     " �����: " + OutDrvFspcAfter/1048576+ " M�")
}

function copy1folder(dirz)
// ���������� ���� �����
{
  //WScript.Echo("copy " + dirz)

  // ���������� ������ ������ � �����������
  if(dirz=="")return(1)
  if(dirz==" ")return(1)
  if(dirz.charAt(0)=="#")return(1)
  if(dirz.charAt(0)==":"){
	l=dirz.length	// �����
	lnz=dirz.indexOf(" ")
	if(lnz==-1)return(1)
	if(l==lnz)return(1)
    namel=dirz.substr(1,lnz)				// 
	pswd=dirz.substr(lnz+1,l)				//
	return(1)	
  }

  // ���� ������ ���������� � \\ (������� �����)
  if(dirz.indexOf("\\\\")!=-1){
    id=1
    l=dirz.length	// �����
    //���� 2-� "\"
    lu=dirz.indexOf("\\",3)
    if(lu!=-1){
      dll=dirz.substr(lu+1,dirz.length+1)	// ��� �������� (������� �����)
      dln=dirz.substr(2,lu-2)				// ��� �����
	  // ��������������� ��� ���.����� � ������� �� FF
	  ldf=dirf.indexOf(".")-1
	  dirff=dirf.substr(1,ldf)
      ff=pref + "\\" +outfp+ fndate+"\\"+dln +"("+dirff+")"
	  try {
	   fso.CreateFolder(ff)
	  }catch(err){
	        writelog("�� ���� ������� " + ff + "  : " + err.number +" : "+ err.description) //��������������� ���������
	        err.clear
	  }     
    }else{
      dll=dirz
      dln=dirz
      writelog("�� ������� ������� ����� ������ ���������:" + dirz ) //��������������� ���������
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
    // ��������� ��� ������ � �������� ������ ������
    if (luu!=-1){
      dln=dln+"_"+dll.substr(0,luu)+"_"+dll.substr(luu+1,dll.length-luu-1)
    }else{
      dln=dln+"_"+dll
    }

  }else{	// ��� ���������
    id=0
    dl  =dirz+ "\\*.*"
    dln =dirz
    ff=pref + "\\" +outfp+ fndate
  }

  if (!fso.FolderExists(dirz)){
    writelog("�� ���� ������� ����� " + dirz ) //��������������� ���������
    return
  }


  // ������� ����� ������ ��� �������,�������� � ��������� �� �������������
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

  arcf = arrc+  ".rar"		//	��� ���.�����
  // ff �������� �������

  // ������ - ������ ���������� �����:
  //    rc=shellobj.Run ("C:\WINNT\system32\cmd.exe /K file.bat", 1,True)
  // ����� �� �� �����, �� ���� �� ������ ...
  
  // ���������� ����� ����� �����
  var rc=-1
  ShellObj.CurrentDirectory=progpath

  wrlog="rar.exe a -ilog"+elgf+" " + a0 + " -m1 -r0 -x@" + pref + "\\exl.txt  " + ff + "\\"+arcf + " \"" + dirz + "\" "
  rc=ShellObj.Run(wrlog ,0,true)

  // ���������� ������ rar.exe
  
  switch(rc){
  case 0: writelog(wrlog+"rc="+rc+" : O'k")
              break
  case 1: writelog(wrlog+"rc="+rc+" : ����������������� �����������")
              break

  case 2: writelog(wrlog+"rc="+rc+" : ��������� ������")
              break
  case 3: writelog(wrlog+"rc="+rc+" : ������ ��(CRC)")
              break
  case 4: writelog(wrlog+"rc="+rc+" : ������������ �����")
              break
  case 5: writelog(wrlog+"rc="+rc+" : ������ ������")
              break
  case 6: writelog(wrlog+"rc="+rc+" : ������ ��������")
              break
  case 7: writelog(wrlog+"rc="+rc+" : ������ ������������")
              break
  case 8: writelog(wrlog+"rc="+rc+" : ������ ������")
              break
  case 9: writelog(wrlog+"rc="+rc+" : ������ ��������")
              break
  case 255: writelog(wrlog+"rc="+rc+" : �������� �������������")
              break
  default: writelog(wrlog+" : ����������� ��� ���������� rc="+rc)
              break
  }
  return(0)
}

function DeleteOldFolders(outfpp,nddel)
// �����e� ������ ������
// outfpp - ��� ������(inc/full)
// nddel  - �� ������� ���� �� �������
{
 var indx3
 var indx
 var tdt = new Date()
  dd=tdt.getDate()
  dd=dd-nddel
  tdt.setDate(dd)
  timdd=tdt.getTime()             // ����� ��� ���������� ����
  // ������ ���� �������� (����������)
  //var tdd=new Date(timaa)
  writelog("�����e� ������ ������ "+outfpp+" �� "+tdt.toLocaleString()+" � ������")

  var f=fso.GetFolder(pref)             // ��� ����� �������
  var ff=new Enumerator(f.SubFolders)   // ����� ��������
  var fc=new Enumerator(f.files)        // � ����� � ���
  var re=new RegExp("^"+outfpp)

  indx=0
  // ��� ���� �����
  for(;!ff.atEnd();ff.moveNext()){
    //�������� �� ����� ����
    s=ff.item().Name
    if(s.match(re)==null)continue
    n=s.indexOf("_")+1
    m=s.indexOf(".")
    da=parseInt(s.substring(n,m))       // ����
    n=s.indexOf(".",m+1)
    mo=parseInt(s.substring(m+1,n))-1   // �����
    ye=parseInt(s.substring(n+1,n+5))   // ���
    var tar=new Date(ye,mo,da)  // �������� ���� ������(������)
    timar=tar.getTime()         // ���������� ���� ������
    // � ����� ������ ���� ��������  ����������
    if(timar>timdd)continue
    // ������� ����������:
    var fcc=fso.getFolder(pref+"\\"+s)  // �������� �����
    var fff=new Enumerator(fcc.files)   // ��������� ������

    // ��� ������� �������� ������ �������
    for(;!fff.atEnd();fff.moveNext()){
      // ������ �.�. �������
      ss=fff.item().Name
      fso.DeleteFile(pref+"\\"+s+"\\"+ss,true)
      writelog(" ���� "+ss+" ������")
      indx++
    }

    fso.DeleteFolder(pref+"\\"+s,true)
    writelog("*����� "+s+" �������")
    indx++
  }

  if (indx==0){
    writelog("*������� ��� �������� ���")
  } else {
    writelog("����� ������� "+indx+" ������/�����")    
  }

  indx=0
  // ��� ���� ������ ����� � ����� ������ ���� �������� �� �������
  for(;!fc.atEnd();fc.moveNext()){
    // �������� �� ����� ����
    s=fc.item().Name
    if(s.match(re)==null)continue
    n=s.indexOf("Log_")
    if(n == -1){
      n=s.indexOf("Err_")
      if(n == -1)continue
    }
    n+=4
    m=s.indexOf(".",n)
    da=parseInt(s.substring(n,m))       // ����
    n=s.indexOf(".",m+1)
    mo=parseInt(s.substring(m+1,n))-1   // �����
    ye=parseInt(s.substring(n+1,n+5))   // ���

    var tar=new Date(ye,mo,da)  // �������� ���� ������(������)
    timar=tar.getTime()         // ����� ��� ���������� ����
    if(timar>timdd)continue
    // � ����� ������ ���� ��������
    // ������� ������� (log@err) ����
    fso.DeleteFile(pref+"\\"+s,true)
    writelog("*���� "+s+" ������")
    indx++
  }

  if (indx==0){
    writelog("*������ ��� �������� ���")
  } else {
        writelog("����� ������� "+indx+" ������")    
  }

 var OutDrvFspcAfter
 OutDrvFspcAfter = OutDrvObj.FreeSpace
 writelog(" ��������� ������������ ����� ��������: " + OutDrvFspcAfter/1048576+ " M�")
}

function writelog(text)
// ������ ������ �����
{
 var logfo
 var ttr=new Date()  // �������� �����
//��������� ���-����
 try {
  logfo = fso.OpenTextFile(logf,8,true)
 }catch(err){
  ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp30  cannot create log file "  + err.number +" : "+ err.description)   // ��������������� ���������
  err.clear
  WScript.Quit()
 }
 logfo.WriteLine(ttr.toLocaleString()+" "+text)
 logfo.close()

}

function testbk()
// �������� ����� Z ������������� ����� � �������� �������� �������� �� ����� ���������� � ��.������
// ���� z: ������ ������ �� �����, �� ������ � ��� ������� ���������� ��������� ������������
{
  if (bkuppc.charAt(0)=="!"){
    createdr=true
    tmp=bkuppc.substr(1,bkuppc.length+1)
    bkuppc=tmp
  }else{
    createdr=false
  }
  //�������� � ����������� ������������ ����� ������ Z:
  if(bkuppc=="Z:\\"){
    try {
      //ntobj.MapNetworkDrive("Z:",bkserv+"backup$",true , "priz", "priz")
      ntobj.MapNetworkDrive("Z:",bkserv+"backup$",true)
    } catch(err){
      if (!(err.number==-2147024811)){//
        //�� ��������� ������ error
        ShellObj.LogEvent(4,"*"+ WScript.Scriptname + " cp1 " + err.number +" : "+ err.description) //��������������� ���������
        err.clear
      }else{
        aaz=1
      }
    }
    driven="Z"
    bkuppc=bkserv+"backup$"
  }else{
    //��� ������ ��������� ��������� � ��������� ������
    if(bkuppc.indexOf(":")!=-1){
      //��������� ���� �� ����� ����
      driven=bkuppc.charAt(0)
      nodrv=0
    }else if(bkuppc.indexOf("\\\\")!=-1){
      //��� ������� ���
      // �������� ����������� ����
      try {
        ntobj.RemoveNetworkDrive("Z:",true,true)
      } catch(err) {
        //�� ��������� ������ 
        ShellObj.LogEvent(4,"*"+ WScript.Scriptname + " cp11 " +"remove error Z:  " + err.number +" : "+ err.description) //��������������� ���������
        err.clear
      }

      l=bkuppc.length
      // ���� ������ �������
      lu=bkuppc.indexOf("\\",3)
      if(lu!=-1){
        dll=bkuppc.substr(lu+1,bkuppc.length+1)
      }else{
        // ��������� ���� ������ (�� ������� ������� �����)
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

      // �������� ���������� ����� � ����� z: ���� ��� �������������
      try {
        ntobj.MapNetworkDrive("Z:",bkuppc.substr(0,la),true)
      } catch(err){
        if (!(err.number==-2147024811)){//
          //�� ��������� ������ error
          ShellObj.LogEvent(4,"*"+ WScript.Scriptname + " cp13 " + err.number +" : "+ err.description) //��������������� ���������
          err.clear
        }else{
          aaz=1
        }
      }
      driven="z"
      nodrv=0
    }else{
      //�� ��������� ������� ��� ��� D
      driven="D"
      nodrv=1
    }
  }

  if(bkuppc.charAt(bkuppc.length-1)!="\\")pre="\\"
  else pre=""

  if(nodrv==1) drv=driven +":\\"
  else  drv=""

  //�������� ������� �������� ��� ����� ��������� ������,���� � �.�.
  pref=drv+bkuppc+pre +ntobj.ComputerName

  if (fso.FolderExists(pref)) {
    //MsgBox "����� ����������"
    if(!fso.FileExists(pref+"\\"+dirf)){
      ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp17 " + "cannot file "+ pref+"\\"+dirf +" open")
      WScript.Quit()
    }
  }else if (createdr){
      //������� ����� � �����
      createus(drv+bkuppc+pre,ntobj.ComputerName)
  }else{
      // ��������� ���� ������(����������� �����)
      ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp18 " + "cannot folder open")
      WScript.Quit()
  }

  // ����� ���� �����������, ���� ����� Z, ��������� �� �� ����� ������ �������
  if(driven=="Z"){
    nozfind()
  }

}

function createus(pr,ne)
// ������� ������� ������������ � ������ ��� ����������� � ��������� ������ dir.txt � exl.txt
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
    // �������� � ����� ������� �������� 
	if (fso.FolderExists("c:\\Users\\")) {
		//����� ���������� - windows7/vista/8
		ts.WriteLine ("c:\\Users\\"+ntobj.UserName)
	}else{
		//����� ��� windows XP
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
//����� ����� ������� ������ ����� ������ ����� Z: �����? �� �� ������ ������ ���� �� ������ ��� ��� ������
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
    //�� ��������� ������
    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp41 "+ err.number +" : "+ err.description) //��������������� ���������
    err.clear
  }
  if(driven!="Z"){
    // ��� Z
    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp42  no Z: drive")                         //��������������� ���������
    WScript.Quit()
  }

}

function whererar()
// ��� �� RAR
{
  path="C:\\Program Files\\WinRAR"

  if (fso.FolderExists(path)) {
	// �� ������� ����������
    //MsgBox "����� ����������"
    if(fso.FileExists(path+"\\rar.exe")){
      progpath=path
      return		

    }
  }else {
	// �� backup-�������
    path=bkserv+"backbin"
    if (fso.FileExists(path+"\\rar.exe")){
      progpath=bkserv+"backbin"
      return
    }else{
        // �� ���� ����� rar
        ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp51 " + "cannot find RAR")
        WScript.Quit()
    }
  }

}

function getprevbksz(outfpp)
// �������� ������ ����������� ������
// outfpp - ��� (Full/Inc) ������
{
  
  tt=""
  err=0
  // ��������� ����� ��������� �����
  try {
    tf=fso.opentextfile(tfile,1)
  } catch(err) {
    //error
    err=1
    ShellObj.LogEvent(4, "*"+ WScript.Scriptname + " cp021 timestamp error "  + err.number +" : "+ err.description)   //��������������� ���������
    err.clear
    //WScript.Quit()
  }
  
  if(err==0){
	 // � ����� ���������������� ������ ������� ���� ��������� �����
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
		lastfull=""							//�� �����
		return -2
	}
    var f1=fso.GetFolder(pref)             	// ��� ����� �������
    var fo=new Enumerator(f1.SubFolders)   	// ����� ��������
    // 1-day,0-month,2-year
    tts=tt[0].split("/")

    //��� ������
    fndat = "_" + tts[1] + "." + tts[0]  + "." + tts[2] + "_"
    nameff=tt[1]+ fndat
    var re=new RegExp("^"+nameff)

    indx=0
    i1=0
    // ��� ���� �������� ���� ����
    for(;!fo.atEnd();fo.moveNext()){
      //�������� ���
      s=fo.item().Name
      if(s.match(re)!=null){
		  // ���� �����!!!
		  i1=1	
		  break
      }
      indx++
    }
    if(i1==0){ // �� �����
      lastfull=""
	   return -3
	}
    //����� �� � ������
	lastfull=pref+"\\"+s    
    var f = fso.GetFolder(lastfull);
    var x = f.Size;	// � ��� �����
    return x
  }
  lastfull=""
  return -1
}

