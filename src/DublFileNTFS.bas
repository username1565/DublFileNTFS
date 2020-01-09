#Include "windows.bi"
#include "win\shellapi.bi"
#Include "window9.bi"
#Include "file.bi"
#Include "vbcompat.bi"
#Include "crt.bi"
Const attrib_mask =  fbNormal Or fbHidden Or fbSystem Or fbDirectory ' для каталогов
Const attrib_mask2 = fbNormal Or fbHidden Or fbSystem ' для файлов
Dim Shared As String folder,file,file2
ReDim Shared spisok(16383) As String ' таблица со списком файлов
Dim Shared razmer() As LongInt ' таблица размеров файла
ReDim Shared sortirovka(0) As Integer ' таблица отсортированного списка файлов (в "0" ячейке размер списка файлов)
Dim As Integer i,k,sovp0:Dim Shared As Integer zam,osh
folder=ShellFolder( "Выбор Каталога Для Обработки", "")' выбираем конечный каталог
If folder="" Then End Else folder=RTrim(folder,"")' если не выбрано, то выходим;  иначе обрезаем слеш вконце

Sub SORT() ' сортировка массива razmer() размера файлов для ускорения поиска 
   Dim As Integer i=1,ind=sortirovka(0):Dim As Byte obmen, revers=1 ' сортировка сверху вниз и снизу вверх для ускорения
   If ind<3 Then Exit Sub ' если сортировать не чего  (меньше не работает)
   Do ' сортируем 
      If razmer(sortirovka(i))<razmer(sortirovka(i+1)) Then Swap sortirovka(i), sortirovka(i+1): obmen=1
      i+=revers ' увеличиваем следующий индек массива в соответствии с напрвлением
      If i=ind OrElse i=0 Then ' если достигли верхнего или нижнего края
       If i=ind Then i-=1 '
       If i=0 Then i=1 
       If obmen=1 Then obmen=0 Else Exit Do ' если обмена за предыдущий проход не было то выходим
       revers*=-1       
      EndIf   
   Loop' сортруется массив razmer() по значению 
End Sub
Sub SPISOK_FAILOV (path As String)'Список файлов каталога   (c подкаталогами)
   Dim As UInteger out_attr,ind_max,ind,sh,pm,i=255:ReDim As String names(255)
   Dim As String fname, dirname = Dir(path  & "\*", attrib_mask, out_attr)
   Do Until Len(dirname) = 0     
    If (dirname <> ".") AndAlso (dirname <> "..") Then ' ignore current and parent directory entries       
       If out_attr And fbDirectory Then 
          If sh>i Then i *=2:ReDim Preserve names(i)
          sh+=1:names(sh)=dirname
       EndIf    
    EndIf
    dirname = Dir(out_attr) ' find next name/attributes
   Loop' каталоги закончились
   For i=1 To sh
      SPISOK_FAILOV (path & "" & names(i))' заходим в подкаталог    
   Next
   fname = Dir(path & "\*", attrib_mask2):ind_max=UBound(spisok):ind=sortirovka(0)' читаем имя файла
    Do Until Len(fname) = 0 ' читаем файлы текущего каталога
          ind+=1:spisok(ind)=path & "" & fname' заполняем таблицу
       If ind=ind_max Then ind_max *=2:ReDim Preserve spisok(ind_max)
    fname = Dir() ' find next name/attributes
    Loop:sortirovka(0)=ind
End Sub
Function DEL_FILE(file As String) As Integer ' возвращает "0" если удачное удаление и ">0" если не удалось
   If  FileExists (file)=0 Then Return 2 ' если такого файла вообще нет, то выходим
   If Kill(file) Then ' файл удалить не получилось
      SetFileAttributes(file, 128)' снимаем все атрибуты файла (возвращает 0 если не удачно) 
      If Kill(file) Then ' если файл удалить снова не удалось
         For i As Integer=1 To 200 ' делаем 200 попыток по 0,1 сек = 20 сек            
            If Kill(file)=0 Then Return 0 ' если удачное удаление то выходим
            Sleep 100  ' делаем задержку 0,1 сек
         Next:Return 1 ' если в течение 20 сек удалить не получилось
      EndIf
   EndIf:Return 0 ' если файл удалён
End Function
Function SRAVNENIE(file1 As String,file2 As String) As Integer ' точное сравнение 2-х файлов
  If Not(FileExists(file1) AndAlso  FileExists(file2)) Then Return 2' если одного из файлов нет
  If file1=file2 Then Return 3' попытка сравнить один и тот же файл   
   Dim As UByte buf1(),buf2()' функция возвращает: 0 -файлы равны; 1 -файлы не равны; 2 -нет файла(ов); 3 -сравнение одного и того же файла
   Dim As UInteger razmer,blok,nf2,nf1=FreeFile   
   Open file1 For Binary Access Read As nf1
   razmer=Lof(nf1):nf2=FreeFile' читаем размер первого файла
   Open file2 For Binary Access Read As nf2      
   If razmer<>Lof(nf2) Then Close nf1,nf2:Return 1 ' если разный размер файлов то выходим
   If razmer>2097151 Then ' если размер файлов больше или равен размеру блока 2097152 байт 
      ReDim buf1(2097151),buf2(2097151):blok=razmer\2097152' Чтение с буфером 2097152 байт      
      For i As UInteger=1 To blok
         Get #nf1,,buf1():Get #nf2,,buf2()
         If memcmp(@buf1(0),@buf2(0), 2097152) Then Close nf1,nf2:Return 1 ' если не равны         
      Next
      razmer=razmer-blok*2097152
      If razmer=0 Then Close nf1,nf2:Return 0 Else ReDim buf1(razmer-1), buf2(razmer-1) ' если нет остатка               
   Else ' если размер файла меньше размера блока
      ReDim buf1(razmer-1),buf2(razmer-1) ' если размер меньше 1 блока         
   EndIf 
   Get #nf1,,buf1():Get #nf2,,buf2():Close nf1,nf2
   If memcmp(@buf1(0),@buf2(0), razmer) Then Return 1 Else Return 0 'если равны возвращаем "0"
End Function 
Function NAMES (name1 As String, name2 As String) As Integer' возвращает "1" если ошибка и "0" если удачное переименование
   For i As Integer=1 To 200 ' делаем 200 попыток по 0,1 сек = 20 сек            
      If Name(name1, name2)=0 Then Return 0 ' если удачное переименование то выходим
      Sleep 100  ' делаем задержку 0,1 сек
   Next:Return 1 ' если в течение 20 сек удалить не получилось   
End Function
Function FILEEXIST (file0 As String) As Integer   ' проверка наличия файла, возвращает "0", если файл есть
   For i As Integer=1 To 200' делаем 200 попыток по 0,1 сек = 20 сек
      If FileExists(file0) Then Return 1
      Sleep 100
   Next:Return 0
End Function
Sub EHO(vb As Integer=0)' отображение информации на экране
   Dim text As String
   If vb=1 Then
      text="Ошибка Создания Файла!"
   ElseIf vb=2 Then
      text="Нет Файлов для Обработки!"
   EndIf
   If vb=0 Then ' вывод результата
      text=Str(zam)
      Select Case Right(Str(zam),1)
         Case "1"
            If zam>10 AndAlso Right(Str(zam),2)="11" Then text &= " Файлов Заменено" Else text &= " Файл Заменён"
         Case "2","3","4"         
            If zam>10 AndAlso ValInt(Right(Str(zam),2))<15 Then text &= " Файлов Заменено" Else text &= " Файла Заменено"
         Case Else
            text &= " Файлов Заменено"
      End Select
      text &= " Ссылками"
      If osh Then text &= Chr(13) & Chr(10) & Str(osh) & " Не Удалось Заменить"
      MessageBox (0,text, "Отчет  DublFileNTFS", MB_OK or MB_ICONINFORMATION)
   Else ' сообщение с ошибкой
      MessageBox (0,text,"", MB_OK or MB_ICONERROR)
   EndIf   
End Sub
Sub DUBL(ByRef nach As Integer,ByRef konec As integer)' обработка блока файлов с одинаковым размером
   Dim As String ishfail,rezfail
   For k As Integer=nach To konec   ' сверяем файл со всеми последующими в блоке
      For i As Integer=k+1 To konec
         If sortirovka(i)=-1 Then Continue For' пропускаем уже созданные ссылки
         If SRAVNENIE(spisok(sortirovka(k)),spisok(sortirovka(i)))=0 Then ' если файлы равны
            file=spisok(sortirovka(k)):file2=spisok(sortirovka(i))
            rezfail=Left(file2,InStrRev(file2,"")) & "$tmp!file$"' выдесяем путь со слешем и добавляем имя временного файла
            If FileExists(rezfail) Then DEL_FILE(rezfail) ' удаляем временный файл, если он случайно остался
            folder=rezfail
            If InStr(file," ") Then ishfail="""" & file & """" Else ishfail=file' если есть пробелы, то заключаем в кавычки
            If InStr(rezfail," ") Then rezfail = """" & rezfail & """" ' если есть пробелы, то заключаем в кавычки            
            ShellExecute (null, null, "fsutil", "hardlink create " & rezfail & " " & ishfail,null, 0)' создаём временный файл-ссылку            
            If FILEEXIST(folder) Then' если временный файл был успешно создан
               If DEL_FILE(file2) Then ' если заменяемый файл не удалось удалить
                  DEL_FILE(folder):osh+=1 ' тогда удаляем временный файл
               Else ' удаление удачное
                  If NAMES(folder,file2) Then DEL_FILE(folder):osh+=1 Else zam+=1:sortirovka(i)=-1 ' блокируем повторную обработку этого файла
               EndIf               
            Else
               If zam=0 Then   EHO 1:End   Else osh+=1   ' если временный файл не удалось создать и не было ни одной замены (FAT)
            EndIf                  
         EndIf      
      Next   
   Next
End Sub

SPISOK_FAILOV folder:k=sortirovka(0) ' получаем список файлов и фиксируем реальное заполнение массивов
If k<2 Then EHO 2:End ' выводим сообщение об отсутствии файлов
ReDim razmer(k)' формируем таблицу размеров нужного диапазона
ReDim Preserve sortirovka(k)' формируем таблицу сортировки нужного диапазона
For i=1 To k
   razmer(i)=FileLen(spisok(i)):sortirovka(i)=i' записываем размер файла в таблицу   
Next:SORT ' сортируем по размеру файлов
For i=1 To k-1
   If razmer(sortirovka(i))<16000 Then Exit For ' пропускаем пустые и мелкие файлы
   If razmer(sortirovka(i))=razmer(sortirovka(i+1)) Then ' совпадение тек. размера с предыдущим
      If sovp0=0 Then   sovp0=i' фиксируем начало совпадения
      If i=k-1 Then DUBL sovp0,i+1' если это последний файл в списке
   Else ' если совпадений нет или они закончились 
      If sovp0 Then ' если есть начало блока совпадений
         DUBL sovp0,i+1:sovp0=0
      EndIf   
   EndIf
Next
EHO
