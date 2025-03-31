

#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	СоздатьДеревоAD();	
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

&НаКлиенте
Процедура ДеревоADФормыВыбор(Элемент, ВыбраннаяСтрока, Поле, СтандартнаяОбработка)
	
	Оповестить("ВыбораГруппыAD_ГруппыПользователейAD", Элемент.ТекущиеДанные.LDAP);
	Закрыть();
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаСервере
Процедура СоздатьДеревоAD()
	
	ДеревоAD = РеквизитФормыВЗначение("ДеревоADФормы");
	ДеревоAD.Строки.Очистить();
	
	DSE = ПолучитьCOMОбъект("LDAP://rootDSE"); 
	LDAP_DNC  = DSE.Get("defaultNamingContext");
	LDAPText = "GC://" + LDAP_DNC;

	// BSLLS:UsingObjectNotAvailableUnix-off Неактуальная проверка работы в UNIX-клиентах
	conn = Новый COMОбъект("ADODB.Connection");
	// BSLLS:UsingObjectNotAvailableUnix-on 
	conn.Provider = "ADSDSOObject";
	conn.Open("ADs Provider");
	
	////////////////////

	// запрос по доменам из глобального каталога
	// BSLLS:UsingObjectNotAvailableUnix-off Неактуальная проверка работы в UNIX-клиентах
	rs =  Новый COMОбъект("ADODB.recordset");
	// BSLLS:UsingObjectNotAvailableUnix-on 
	rs.Open("<" + LDAPText + ">;(objectClass=domain);ADsPath, Name, distinguishedName;subtree", conn, 0, 1);  //source,actconn,cursortyp,locktyp,opt

	Если rs.RecordCount > 0 Тогда 
		Пока Не rs.EOF Цикл  
			
			НоваяСтрока = ДеревоAD.Строки.Добавить();
			
			НоваяСтрока.Name			 = rs.Fields("Name").Value;
			НоваяСтрока.LDAP		 = rs.Fields("distinguishedName").Value; 
			НоваяСтрока.Path		 = rs.Fields("ADsPath").Value;  //берем путь "как есть", в т.ч. GC
			
			ДобавитьВетвьВДеревоAD(conn, НоваяСтрока);
			
			Попытка
				rs.MoveNext();
			Исключение
				ОбщегоНазначения.СообщитьПользователю(НСтр("ru = 'Превышен допустимый размер получаемых данных.'"));
				Прервать;
			КонецПопытки;
			
		КонецЦикла;
	КонецЕсли;
	
	rs.Close();
	rs = Неопределено;

	conn.Close();
	conn = Неопределено;
		
	ЗначениеВДанныеФормы(ДеревоAD, ДеревоADФормы);	
	
КонецПроцедуры  //ЗаполнитьТаблицуПользователей

&НаСервере
Процедура ДобавитьВетвьВДеревоAD(conn, СтрокаДерева)
	
	// BSLLS:UsingObjectNotAvailableUnix-off Неактуальная проверка работы в UNIX-клиентах
	rs =  Новый COMОбъект("ADODB.recordset");
	// BSLLS:UsingObjectNotAvailableUnix-on 
	rs.Open("<" + СтрокаДерева.Path + ">;(objectClass=organizationalUnit);ADsPath, Name, distinguishedName;onelevel", conn, 0, 1);  //source,actconn,cursortyp,locktyp,opt
	
	Если rs.RecordCount > 0 Тогда 
		
		Пока Не rs.EOF Цикл
			
			НоваяСтрока = СтрокаДерева.Строки.Добавить();
			
			НоваяСтрока.Path		 = rs.Fields("ADsPath").Value;
			НоваяСтрока.Name		 = rs.Fields("Name").Value;
			НоваяСтрока.LDAP		 = rs.Fields("distinguishedName").Value; 
			
			ДобавитьВетвьВДеревоAD(conn, НоваяСтрока);
			
			Попытка
				rs.MoveNext();
			Исключение
				ОбщегоНазначения.СообщитьПользователю(НСтр("ru = 'Превышен допустимый размер получаемых данных.'"));
				Прервать;
			КонецПопытки;
			
		КонецЦикла;
		
		СтрокаДерева.Строки.Сортировать("Name");
		
	КонецЕсли;
	
	rs.Close();
	rs = Неопределено;

КонецПроцедуры

#КонецОбласти

