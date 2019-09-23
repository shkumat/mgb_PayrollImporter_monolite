// Версия 2.01 от 20 августа 2019г. Загрузка зарплатных ведомостей (dif,dbf) в Скрудж.
//  
// в версии 2 добавилена обратботка IBAN 
//
using	MyTypes;
using	__	=	MyTypes.CCommon ;
using	money	=	System.Decimal ;

public	class	PayrollImporter {
	//  ---------------------------------------------------------------
	//  Функцию Main нужно пометить атрибутом [STAThread], чтоб работал OpenFileBox
	[System.STAThread]
	static void Main()  { //FOLD00
		const	bool	DEBUG		=	false		;
		bool			UseErcRobot	=	false		;	// подключаться ли к серверу под логином ErcRobot
		const	string	ROBOT_LOGIN	=	"ErcRobot"	;
		const	string	ROBOT_PWD	=	"35162987"	;
		const	string	TASK_CODE	=	"OpenGate"	;
		CPayRoll		PayRoll						;
		CCommand		Command						;
		CConnection		Connection					;
		COpengateConfig	OgConfig					;
		CScrooge2Config	Scrooge2Config				;
		byte			SavedColor					;
		int				ErcDate		=	0
		,				SeanceNum	=	0			;
		string			TmpDir		=	CAbc.EMPTY
		,				StatDir		=	CAbc.EMPTY
		,				BankCode	=	"351629"
		,				FileName	=	CAbc.EMPTY
		,				TodayDir	=	CAbc.EMPTY
		,				InputDir	=	CAbc.EMPTY
		,				DataBase	=	CAbc.EMPTY
		,				ScroogeDir	=	CAbc.EMPTY
		,				SettingsDir	=	CAbc.EMPTY
		,				ServerName	=	CAbc.EMPTY
		,				AboutError	=	CAbc.EMPTY
		,				LogFileName	=	CAbc.EMPTY
		,				TmpFileName	=	CAbc.EMPTY
		,				CleanFileName=	CAbc.EMPTY
		,				ConnectionString=CAbc.EMPTY	;
		CConsole.Clear();
		__.Print( CAbc.EMPTY,"  Загрузка зарплатных ведомостей в `Скрудж`. Версия 2.01 от 21.08.2019г." , CAbc.EMPTY );
		System.Console.Title="Загрузка в `Скрудж` зарплатных ведомостей";
		__.DeleteOldTempDirs("??????" , __.Today() - 1 );
		if	( DEBUG ) {
			FileName	=	"D:\\WorkShop\\zkm.dif";
			Err.LogTo("D:\\WorkShop\\zkm.log");
		}
		else
			if	( __.ParamCount() > 1 )
				for	( int i = 1 ; i < __.ParamCount() ; i++ )
					if	( CAbc.ParamStr[ i ].Trim().ToUpper() == "/R" ) {
						UseErcRobot	= true;
						System.Console.Title = System.Console.Title + " * ";
					}
					else
						FileName	=	CAbc.ParamStr[ i ].Trim();
		if	( __.IsEmpty( FileName ) ) {
			__.Print( " Не указано имя файла для обработки ! " );
			__.Print( "  Формат запуска :   PayrollImporter.exe  <FileName>  [/R]" );
			__.Print( "  Пример         :   PayrollImporter.exe  * " );
			return;
		}
		// -----------------------------------------------------
		// Вычитываем настройки "Скрудж-2"
		Scrooge2Config	= new	CScrooge2Config();
		if	(!Scrooge2Config.IsValid) {
			__.Print( Scrooge2Config.ErrInfo ) ;
			return	;
		}
		ScroogeDir	=	(string)Scrooge2Config["Root"]		;
		SettingsDir	=	(string)Scrooge2Config["Common"]	;
		ServerName	=	(string)Scrooge2Config["Server"]	;
		DataBase	=	(string)Scrooge2Config["DataBase"]	;
		if	( ScroogeDir == null ) {
			__.Print("  Не найдена переменная `Root` в настройках `Скрудж-2` ");
			return;
		}
		if	( ServerName == null ) {
			__.Print("  Не найдена переменная `Server` в настройках `Скрудж-2` ");
			return;
		}
		if	( DataBase == null ) {
			__.Print("  Не найдена переменная `Database` в настройках `Скрудж-2` ");
			return;
		}
                ScroogeDir	=	ScroogeDir.Trim()	;
                if	( SettingsDir != null )
                	SettingsDir	=	ScroogeDir + "\\" + SettingsDir ;
		ServerName	=	ServerName.Trim()	;
		DataBase	=	DataBase.Trim()		;
		if	( DEBUG )
			__.Print("  Беру настройки `Скрудж-2` здесь :  " + ScroogeDir );
		if	( FileName == "*" )
			FileName	=	SelectFileNameGUI( SettingsDir );
		if	( FileName == null )
			return;
		if	( __.IsEmpty( FileName ) )
			return;
		if	( ! __.FileExists( FileName ) ) {
			__.Print( " Не найден указанный файл ! " , "" , "Для выхода нажмите ENTER.");
			CConsole.WaitForEscOrEnter();
			return;
		}
		CleanFileName	=	__.GetFileName( FileName );
		PayRoll			= new	CPayRoll();
		if	( ! PayRoll.Open( FileName ) ) {
			__.Print( "Ошибка открытия исходного файла !" , "" , "Для выхода нажмите ENTER.");
			PayRoll.Close();
			CConsole.WaitForEscOrEnter();
			return;
		}
		if	( ! PayRoll.Preview() ) {
			PayRoll.Close();
			return;
		}
		// -----------------------------------------------------
		// Подключаемся к базе данных
		ConnectionString	=	"Server="		+	ServerName
							+	";Database="	+	DataBase ;
		if	( UseErcRobot )
			ConnectionString	+=	";UID=" + ROBOT_LOGIN + ";PWD=" + ROBOT_PWD + ";" ;
		else
			ConnectionString	+=	";Integrated Security=TRUE;" ;
		Connection		= new	CConnection( ConnectionString );
		if	( Connection.IsOpen() ) {
			if	( DEBUG )
				__.Print("  Сервер        :  " + ServerName );
		}
		else {
			__.Print( "","  Ошибка подключения к серверу !" );
			Connection.Close();
			PayRoll.Close();
			__.Print( "" , "Для выхода нажмите ENTER.");
			CConsole.WaitForEscOrEnter();
			return;
		}
		Command			= new	CCommand(Connection) ;
		if	( Command.IsOpen() ) {
			if	( DEBUG )
				__.Print("  База данных   :  " + DataBase );
		}
		else {
			__.Print( "","  Ошибка подключения к базе данных !" );
			Command.Close();
			Connection.Close();
	        	PayRoll.Close();
			__.Print( "" , "Для выхода нажмите ENTER.");
			CConsole.WaitForEscOrEnter();
			return;
		}
		System.Console.Title=System.Console.Title+" |   "+ServerName+"."+DataBase	;
		// -----------------------------------------------------
		//  Вычитываем из БД информацию про МФО и дебет-счет
		BankCode	=	( string ) __.IsNull( Command.GetScalar( " select Code from dbo.vMega_Common_MyBankInfo with ( NoLock ) " ) , CAbc.EMPTY );
		if	( BankCode == null )
			BankCode = "351629" ;
		else
			if	( __.IsEmpty( BankCode ) )
				BankCode="351629" ;
		if	( ! PayRoll.GetDebitInfo( Command ) ) {
			Command.Close();
			Connection.Close();
	        	PayRoll.Close();
			return	;
		}
		// - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		// считываем настройки шлюза в ЕРЦ
		ErcDate			=	( int ) __.IsNull( Command.GetScalar( " exec  dbo.pMega_OpenGate_Days;7 " ) , (int) 0 );
		if	( ErcDate < 1 ) {
			__.Print( " Ошибка определения даты текущего рабочего дня. " );
			Command.Close();
			Connection.Close();
			PayRoll.Close();
			__.Print( "" , "Для выхода нажмите ENTER.");
			CConsole.WaitForEscOrEnter();
			return	;
		}
		OgConfig	= new	COpengateConfig();
		OgConfig.Open( ErcDate );
		if	( ! OgConfig.IsValid() ) {
			__.Print( "  Ошибка чтения настроек программы из " + OgConfig.Config_FileName() );
			__.Print( OgConfig.ErrInfo())		;
			Command.Close();
			Connection.Close();
			PayRoll.Close();
			__.Print( "" , "Для выхода нажмите ENTER.");
			CConsole.WaitForEscOrEnter();
			return;
		}
		SeanceNum	=	( int ) __.IsNull( Command.GetScalar(" exec dbo.pMega_OpenGate_Days;4  @TaskCode='" + TASK_CODE + "',@ParamCode='NumSeance' ") , (int) 0 );
		if	( SeanceNum < 1 ) {
			__.Print( " Ошибка определения номера сеанса " );
			Command.Close();
			Connection.Close();
			PayRoll.Close();
			__.Print( "" , "Для выхода нажмите ENTER.");
			CConsole.WaitForEscOrEnter();
			return	;
		}
		TodayDir	=	(string)OgConfig.TodayDir()		;
		TmpDir		=	(string)OgConfig.TmpDir()		;
		StatDir		=	(string)OgConfig.StatDir()		;
		if ( (TodayDir == null) || (InputDir == null) ) {
			__.Print( "  Ошибка чтения настроек программы из " + OgConfig.Config_FileName() );
			Command.Close();
			Connection.Close();
			PayRoll.Close();
			__.Print( "" , "Для выхода нажмите ENTER.");
			CConsole.WaitForEscOrEnter();
			return;
		}
		TodayDir	=	TodayDir.Trim() ;
		StatDir		=	StatDir.Trim();
		TmpDir		=	TmpDir.Trim();
		if	( ! __.DirExists( TodayDir ) )
			__.MkDir( TodayDir );
		if	( ! __.DirExists( StatDir ) )
			__.MkDir( StatDir );
		if	( ! __.DirExists( TmpDir ) )
			__.MkDir( TmpDir );
		if	( ! __.SaveText( StatDir + "\\" + "test.dat" , "test.dat" , CAbc.CHARSET_DOS ) ) {
			__.Print( " Ошибка записи в каталог " + StatDir );
			Command.Close();
			Connection.Close();
			PayRoll.Close();
			__.Print( "" , "Для выхода нажмите ENTER.");
			CConsole.WaitForEscOrEnter();
			return	;
		}
		__.DeleteFile(StatDir + "\\" + "test.dat");
		LogFileName	=	OgConfig.LogDir() + "\\SEANS" + SeanceNum.ToString("000")  + ".TXT";
		Err.LogTo( LogFileName );
		__.AppendText( LogFileName , __.Now() + "   " + __.Upper(__.GetUserName()) + "  загружает файл " + CleanFileName + CAbc.CRLF , CAbc.CHARSET_WINDOWS );
		TmpFileName		=	TodayDir + CAbc.SLASH + CleanFileName	;
		if	( ! __.FileExists( TmpFileName ) )
			__.CopyFile( FileName , TmpFileName ) ;
		if	( ! __.FileExists( TmpFileName ) ) {
			__.Print( " Ошибка записи файла " + TmpFileName );
			Command.Close();
			Connection.Close();
			PayRoll.Close();
			__.Print( "" , "Для выхода нажмите ENTER.");
			CConsole.WaitForEscOrEnter();
			return ;
		}
		TmpFileName		=	TmpDir + CAbc.SLASH
					+	__.Right( "0" + __.Hour(__.Clock()).ToString() , 2 )
					+	__.Right( "0" + __.Minute( __.Clock()).ToString() , 2 )
					+	__.Right( "0" + __.Second( __.Clock()).ToString() , 2 )	;
		if( ! __.DirExists( TmpFileName ) )
			__.MkDir( TmpFileName )	;
		PayRoll.CopyTempFile( TmpFileName );
		TmpFileName		=	TmpFileName + CAbc.SLASH + CleanFileName	;
		PayRoll.CopyTempFile( TmpFileName );
		if	( __.FileExists( TmpFileName ) )
			__.DeleteFile( TmpFileName )	;
		if	( __.FileExists( TmpFileName ) ) {
			__.Print("Ошибка удаления файла ",TmpFileName)	;
			Command.Close();
			Connection.Close();
			PayRoll.Close();
			__.Print( "" , "Для выхода нажмите ENTER.");
			CConsole.WaitForEscOrEnter();
			return	;
		}
		__.CopyFile( FileName , TmpFileName )	;
		if	( DEBUG )
			__.Print("  Беру настройки шлюза здесь :  " + OgConfig.Config_FileName() );
		// -----------------------------------------------------
		//  Проверяем пачку
		AboutError	=	PayRoll.CheckAll( Command , BankCode ) ;
		if	( AboutError != CAbc.EMPTY ) {
			__.Print( AboutError );
			__.AppendText( LogFileName ,  CAbc.CRLF + AboutError + CAbc.CRLF , CAbc.CHARSET_WINDOWS );
			SavedColor		=	CConsole.BoxColor ;
			CConsole.BoxColor	=	CConsole.RED*16 + CConsole.WHITE	;
			if	( ! CConsole.GetBoxChoice(	" При проверке файла обнаружены ошибки !"
							,	" Для отмены загрузки нажмите ESC . "
							,	" Для продолжения - ENTER ."
							)
				) {
				__.AppendText( LogFileName ,  CAbc.CRLF + __.Now() + "  загрузка отменена. " + CAbc.CRLF , CAbc.CHARSET_WINDOWS );
				CConsole.BoxColor	=	SavedColor ;
				Command.Close();
				Connection.Close();
				PayRoll.Close();
				return;
			}
			CConsole.BoxColor	=	SavedColor ;
		}
		else	{
			if	( ! CConsole.GetBoxChoice(	" При проверке файла ошибок не найдено."
							,	" Для загрузки нажмите ENTER . "
							,	" Для выхода - ESC."
							)
				) {
				__.AppendText( LogFileName ,  CAbc.CRLF + __.Now() + "  загрузка отменена. " + CAbc.CRLF , CAbc.CHARSET_WINDOWS );
				Command.Close();
				Connection.Close();
				PayRoll.Close();
				return;
			}
		}
		// -----------------------------------------------------
		//  Загружаем пачку
		if	( ! PayRoll.InsertAll( Command , BankCode ) ) {
			SavedColor		=	CConsole.BoxColor ;
			CConsole.BoxColor	=	CConsole.RED*16 + CConsole.WHITE	;
			CConsole.GetBoxChoice( CAbc.EMPTY ,"  При загрузке файла были ошибки !" , CAbc.EMPTY )  ;
			CConsole.BoxColor	=	SavedColor ;
		}
		// -----------------------------------------------------
		Command.Close();
		Connection.Close();
		PayRoll.Close();
		__.AppendText( LogFileName ,  CAbc.CRLF + __.Now() + "  загрузка завершена. " + CAbc.CRLF , CAbc.CHARSET_WINDOWS );
	}//FOLD00
	//  -------------------------------------------------------------
	//  Получить имя файла с помощью графической панели открытия файла
	static string SelectFileNameGUI( string SettingsPath ) {//fold00
		string		Result		=	CAbc.EMPTY;
		string		SettingsFileName=	null;
		if	( SettingsPath != null )
			if	( SettingsPath.Trim().Length > 0 ) {
		      		SettingsFileName	=	SettingsPath.Trim() + "\\" + CCommon.GetUserName() + ".ldr";
		      		if	( CCommon.FileExists( SettingsFileName ) )
					Result		=	CCommon.LoadText(  SettingsFileName , CAbc.CHARSET_WINDOWS );
				if	( Result == null )
				        Result	=	CAbc.EMPTY;
			}
		Result	=	Result.Trim();
		Result	=	__.OpenFileBox(
					"Укажите файл для обработки"
				,	Result
				,	"з/п ведомости (*.dif,*.dbf)|*.d?f"
			);
		if	( Result == null )
			return	CAbc.EMPTY;
		Result		=	Result.Trim();
                if	( __.IsEmpty( Result ) )
			return	Result;
		if	( SettingsFileName != null )
			CCommon.SaveText( SettingsFileName , __.GetDirName( Result ) , CAbc.CHARSET_WINDOWS ) ;
		return	Result;
	}//FOLD00
}

//  Вычитывателя конфигурации универсального шлюза//fold00
public	class	COpengateConfig	:	CErcConfig {
	public	override string StatDir() {
		return TodayDir() + "\\STA\\";
	}
        public	override string Config_FileName() {
        	return	"EXE\\GLOBAL.FIL";
        }
	public	override string	TodayDir() {
		string TmpS = __.DtoC(Erc_Date);
		return CfgFile["DaysDir"] + "\\" + TmpS.Substring(2, 6) + "\\";
	}
}//FOLD00

// Зарплатная ведомость
public	class	CPayRoll {
	const	int		FLD_DEBITMFO	=	0;
	const	int		FLD_DEBITACC	=	1;
	const	int		FLD_DEBITNAME	=	2;
	const	int		FLD_DEBITSTATE	=	3;
	const	int		FLD_CREDITMFO	=	4;
	const	int		FLD_CREDITACC	=	5;
	const	int		FLD_CREDITNAME	=	6;
	const	int		FLD_CREDITSTATE	=	7;
	const	int		FLD_SUMA		=	8;
	const	int		FLD_CURRENCY	=	9;
	const	int		FLD_PURPOSE		=	10;
	const	int		FLD_CODE		=	11;
	const	int		TOTAL_FIELDS	=	12;
	const	string	TASK_CODE		=	"OpenGate"	;
	readonly char	TAB				=	System.Convert.ToChar(9) ;
	readonly string	USER_NAME		=	__.Upper( __.GetUserName() ) ;
	readonly int	WIDTH			=	System.Console.WindowWidth - 1 ;
	readonly int	HEIGHT			=	System.Console.WindowHeight - 1 ;
	readonly string	ModelFileName	=	__.GetTempDir() + "\\" + "AMaker.mod" ;
	string[]		FieldValues		= new string[ TOTAL_FIELDS ];
	int[]			FieldFromFile	= new int[ TOTAL_FIELDS ];
	string		CleanFileName		=	"(unknown)"	; // имя файла без путей
	int		FileType	=	0		  // 1=dbf,2=dif,3=csv,4=mod
	,		DayDate		=	__.Today()
	,		CharSet		=	CAbc.CHARSET_WINDOWS;
	string		FileName	=	CAbc.EMPTY	;
	money		TotalSum	=	0	;
	int		TotalLines	=	0	;
	int		TotalColumns	=	0	;
	bool		NeedDelSrcFile	=	false	;
	CCfgFile	ModelFile	=	null	;
	IFileOfColumnsReader Reader	=	null	;

	public	void	Close() {//fold00
		if	( Reader != null )
			Reader.Close();
		if	( NeedDelSrcFile )
			if	( __.FileExists( FileName ) )
				__.DeleteFile( FileName ) ;
	}//FOLD00

	public	bool	Open( string File_Name ) {//fold00
		if	( File_Name == null )
			return	false;
		if	( __.IsEmpty( File_Name ) )
			return	false;
		FileName	=	File_Name;
		CleanFileName	=	__.GetFileName( FileName );
		switch	( __.GetExtension( CleanFileName ).ToUpper().Trim() ) {
			case	".DBF" : {
				FileType	=	1;
				CharSet		=	CAbc.CHARSET_DOS;
				break;
			}
			case	".DIF" : {
				FileType	=	2;
				break;
			}
			case	".CSV" : {
				FileType	=	3;
				break;
			}
			case	".MOD" : {
				FileType	=	4;
				break;
			}
			default	: {
				FileType	=	0;
				break;
			}
		}
		if	( FileType == 0 )
			return false ;
		if	( FileType == 4 ) {	// mod
			ShowModel();
			return false ;
		}
		if	( FileType == 2 ) {	// dif
			if	( ! ConvertDifToCsv() )
				return	false;
			FileType = 3 ;		// csv
		}
		for	( int I = 0  ; I < TOTAL_FIELDS ; I++  ) {
			FieldValues[ I ]	=	CAbc.EMPTY ;
			FieldFromFile[ I ]	=	0 ;
		}
		FieldValues  [ FLD_CODE         ] =	"1" /* __.Clock().ToString().Trim() */ ;
		FieldValues  [ FLD_CURRENCY     ] =	"980";
		LoadModelInfo();
		if	( FileType == 1 )
			Reader		= new	CDbfReader();
		else
			Reader		= new	CCsvReader();
		if	( ! Reader.Open( FileName , CharSet ) ) {
			Reader.Close();
			return	false;
		}
		Reader.Close();
		return	true;
	}//FOLD00

	//  Подготовка столбцов к обработке
	void		PrepareFields() {//fold00
		for	( int I = 0 ; I < TOTAL_FIELDS  ; I++ )
			if	( FieldFromFile[ I ] > 0 )
				FieldValues[ I ] = Reader[ FieldFromFile[ I ] ] ;
		money	Suma	=	__.CCur( FieldValues[ FLD_SUMA ].Trim().Replace(" ",CAbc.EMPTY) ) * 100 ;
		FieldValues[ FLD_SUMA ]	= __.CLng( __.Trunc( Suma ) ).ToString() ;
		FieldValues[ FLD_CODE ]	= ( __.CInt( "0" + FieldValues[ FLD_CODE ].Trim() )  + 1 ).ToString() ;
		FieldValues[ FLD_PURPOSE ]	=	__.FixUkrI(	FieldValues[ FLD_PURPOSE ]	) ;
		FieldValues[ FLD_DEBITNAME ]	=	__.FixUkrI(     FieldValues[ FLD_DEBITNAME ]	) ;
		FieldValues[ FLD_CREDITNAME ]	=	__.FixUkrI(     FieldValues[ FLD_CREDITNAME ]	) ;
	}//FOLD00

	//  Загрузка из шаблона информации о дебетовом счете
	public	void	LoadModelInfo() {//fold00
		if	( ! __.FileExists( ModelFileName ) )
			return;
		ModelFile      = new   CCfgFile( ModelFileName ) ;
		string	DebitMoniker    =       (string) ModelFile["AccountA_Text"];
		if	( DebitMoniker != null )
			if	( ! __.IsEmpty( DebitMoniker ) )
				FieldValues[ FLD_DEBITACC ] = DebitMoniker;
		string	Purpose    =       (string) ModelFile["Argument_Text"];
		if	( Purpose != null )
			if	( ! __.IsEmpty( Purpose ) )
				FieldValues[ FLD_PURPOSE ] = __.FixUkrI( Purpose );
		string	Encoding    =       (string) ModelFile["Encoding"];
		if	( Encoding != null )
			if	( Encoding.ToUpper().Trim() == "1251" )
				CharSet = CAbc.CHARSET_WINDOWS ;
			else
				CharSet = CAbc.CHARSET_DOS ;
	}//FOLD00

	//  Вставка строк входного файла в БД
	public	bool	InsertAll( CCommand Command , string BankCode ) {//fold00
		string	CmdText		=	CAbc.EMPTY	
		,	DebitIBAN	=	CAbc.EMPTY
		,	CreditIBAN	=	CAbc.EMPTY		
		,	DebitMoniker	=	CAbc.EMPTY
		,	CreditMoniker	=	CAbc.EMPTY	;
		int	LineNum		=	0		;
		bool	Result		=	true		;
		FieldValues[ FLD_DEBITMFO ] = BankCode ;
		FieldValues[ FLD_CREDITMFO ] = BankCode ;
		if	( ! Reader.Open( FileName , CharSet ) ) {
			Reader.Close();
			return	false;
		}
		while	( Reader.Read() ) {
			LineNum++	;
			PrepareFields() ;
			DebitMoniker	=	FieldValues[ FLD_DEBITACC    ].Replace("'","`").Trim();
			CreditMoniker	=	FieldValues[ FLD_CREDITACC   ].Replace("'","`").Trim();
			if	( DebitMoniker[0] > '9' ) {
				DebitIBAN	=	DebitMoniker;
				DebitMoniker	=	CAbc.EMPTY;
			}
			if	( CreditMoniker[0] > '9' ) {
				CreditIBAN	=	CreditMoniker;
				CreditMoniker	=	CAbc.EMPTY;
			}
			CConsole.ShowBox(CAbc.EMPTY," Загружается строка" + __.StrI( LineNum , 5 ) + " " ,CAbc.EMPTY) ;
			CmdText		=	"exec  dbo.pMega_OpenGate_AddPalvis "
					+	" @TaskCode     = '" + TASK_CODE + "'"
					+	",@DayDate	= "  + DayDate.ToString()
					+	",@BranchCode   = ''"
					+	",@FileName     = '" + CleanFileName + "'"
					+	",@LineNum      =  " + LineNum.ToString()
					+	",@SourceCode   = '" + FieldValues[ FLD_DEBITMFO    ].Replace("'","`").Trim()	+ "'"
					+	",@DebitMoniker = '" + DebitMoniker + "'"
					+	",@DebitState   = '" + FieldValues[ FLD_DEBITSTATE  ].Replace("'","`").Trim() + "'"
					+	",@DebitIBAN    = '" + DebitIBAN + "'"
					+	",@DebitName    = '" + FieldValues[ FLD_DEBITNAME   ].Replace("'","`").Trim()	+ "'"
					+	",@TargetCode   = '" + FieldValues[ FLD_CREDITMFO   ].Replace("'","`").Trim()	+ "'"
					+	",@CreditMoniker= '" + CreditMoniker + "'"
					+	",@CreditState  = '" + FieldValues[ FLD_CREDITSTATE ].Replace("'","`").Trim()	+ "'"
					+	",@CreditIBAN   = '" + CreditIBAN + "'"
					+	",@CreditName   = '" + FieldValues[ FLD_CREDITNAME  ].Replace("'","`").Trim()	+ "'"
					+	",@CrncyAmount  =  " + FieldValues[ FLD_SUMA        ].Replace(" ",CAbc.EMPTY).Trim()
					+	",@CurrencyId   =  " + FieldValues[ FLD_CURRENCY    ].Replace("'","`").Trim()
					+	",@Purpose      = '" + FieldValues[ FLD_PURPOSE     ].Replace("'","`").Replace("?","i").Trim()	+ "'"
					+	",@Code         = '" + FieldValues[ FLD_CODE        ].Replace("'","`").Trim()	+ "'"
					+	",@Ctrls        = ''"
					+	",@UserName     = '" + USER_NAME + "'"
					;
			if	( ! Command.Execute( CmdText ) )
				Result	=	false;
		}
		CConsole.Clear()	;
		CConsole.ShowBox(CAbc.EMPTY," Подождите..." ,CAbc.EMPTY) ;
		Command.Execute("  exec pMega_OpenGate_PayRoll;2 "
			+	"  @FileName='"  + CleanFileName + "'"
			+	", @DayDate=" + DayDate.ToString()
			) ;
		CConsole.Clear();
		Reader.Close();
		return	Result;
	}//FOLD00

	//  Проверка всех строк входного файла
	public	string	CheckAll( CCommand Command , string BankCode ) {//fold00
		string	Result		=	CAbc.EMPTY
		,	CmdText		=	CAbc.EMPTY
		,	SavedCode	=	CAbc.EMPTY
		,	AboutError	=	CAbc.EMPTY
		,	DebitIBAN	=	CAbc.EMPTY
		,	CreditIBAN	=	CAbc.EMPTY		
		,	DebitMoniker	=	CAbc.EMPTY
		,	CreditMoniker	=	CAbc.EMPTY	;
		bool	HaveError	=	false		;
		int	LineNum		=	0		;
		FieldValues[ FLD_DEBITMFO ] = BankCode ;
		FieldValues[ FLD_CREDITMFO ] = BankCode ;
		if	( ! Reader.Open( FileName , CharSet ) ) {
			Reader.Close();
			return	"Ошибка чтения исходного файла !";
		}
		SavedCode	=	FieldValues[ FLD_CODE ];
		while	( Reader.Read() ) {
			LineNum++	;
			PrepareFields();
			DebitMoniker	=	FieldValues[ FLD_DEBITACC    ].Replace("'","`").Trim();
			CreditMoniker	=	FieldValues[ FLD_CREDITACC   ].Replace("'","`").Trim();
			if	( DebitMoniker[0] > '9' ) {
				DebitIBAN	=	DebitMoniker;
				DebitMoniker	=	CAbc.EMPTY;
			}
			if	( CreditMoniker[0] > '9' ) {
				CreditIBAN	=	CreditMoniker;
				CreditMoniker	=	CAbc.EMPTY;
			}
			CConsole.ShowBox(CAbc.EMPTY," Проверяется строка" + __.StrI( LineNum , 5 ) + " " ,CAbc.EMPTY)	;
			CmdText		=	"exec  dbo.pMega_OpenGate_CheckPalvis "
					+	" @Code         = '" + FieldValues[ FLD_CODE        ].Replace("'","`").Trim()	+ "'"
					+	",@Ctrls        = ''"
					+	",@SourceCode   = '" + FieldValues[ FLD_DEBITMFO    ].Replace("'","`").Trim()	+ "'"
					+	",@DebitMoniker = '" + DebitMoniker + "'"
					+	",@DebitState   = '" + FieldValues[ FLD_DEBITSTATE  ].Replace("'","`").Trim() + "'"
					+	",@DebitIBAN    = '" + DebitIBAN + "'"
					+	",@TargetCode   = '" + FieldValues[ FLD_CREDITMFO   ].Replace("'","`").Trim()	+ "'"
					+	",@CreditMoniker= '" + CreditMoniker + "'"
					+	",@CreditState  = '" + FieldValues[ FLD_CREDITSTATE ].Replace("'","`").Trim()	+ "'"
					+	",@CreditIBAN   = '" + CreditIBAN + "'"
					+	",@CrncyAmount  =  " + FieldValues[ FLD_SUMA        ].Replace(" ",CAbc.EMPTY).Trim()
					+	",@CurrencyId   =  " + FieldValues[ FLD_CURRENCY    ].Replace("'","`").Trim()
					+	",@UserName     = '" + USER_NAME + "'"
					;
           		AboutError	=	(string) __.IsNull( Command.GetScalar( CmdText ) , CAbc.EMPTY ) ;
			if	( __.IsEmpty( FieldValues[ FLD_PURPOSE ].Trim() ) )
				AboutError	+=	" Не заполнено назначение платежа ;" ;
			if	( __.IsEmpty( FieldValues[ FLD_DEBITNAME ].Trim() ) )
				AboutError	+=	" Не заполнено название дб. счета ;" ;
			if	( __.IsEmpty( FieldValues[ FLD_CREDITNAME ].Trim() ) )
				AboutError	+=	" Не заполнено название кт. счета ;" ;
			if	( AboutError != null )
				if	( ( AboutError.Trim() != "" ) ) {
						HaveError	=	true;
						Result		+=	" Ошибка в строке " + LineNum.ToString() +" : " + AboutError.Trim()  + CAbc.CRLF  ;
				}
		}
		FieldValues[ FLD_CODE ]	=	SavedCode;
		CConsole.Clear();
		byte	SavedColor		=	CConsole.BoxColor;
		if	( ( ( int ) __.IsNull( Command.GetScalar( "exec dbo.pMega_OpenGate_CheckPalvis;2 @TaskCode='" + TASK_CODE + "',@FileName='" + CleanFileName + "'" ) , (int) 0 ) ) > 0 ) {
			CConsole.BoxColor	=	CConsole.RED*16 + CConsole.WHITE	;
			CConsole.GetBoxChoice( "Внимание ! Файл " + CleanFileName + " сегодня уже загружался !" , "" ,"Для выхода нажмите ENTER.") ;
			CConsole.BoxColor	=	SavedColor	;
			Result		+=	"Файл " + CleanFileName + " сегодня уже загружался !" + CAbc.CRLF ;
		}
		CConsole.BoxColor	=	SavedColor	;
		Reader.Close();
		return	Result;
	}//FOLD00

	//  Разпознание колонок входного файла
	public	bool	Preview() {//fold00
		int	Choice			=	0 ;
		const	int	MAX_COL		=	100;	// Максимальное кол-во столбцов  в файле
		string[]Lines		= new	string[ HEIGHT ] ;
		int[]	Sizes		= new	int[ MAX_COL ] ;
		if	( ! Reader.Open( FileName , CharSet ) ) {
			Reader.Close();
			return	false;
		}
		TotalLines	=	Reader.Count;
		int	I,LineNum	=	0;
		for	( I = 0 ; I < MAX_COL ; I ++ )
			Sizes[ I ] = 0;
		for	( I = 0 ; I < HEIGHT ; I ++ )
			Lines[ I ] = CAbc.EMPTY ;
		while	( Reader.Read() ) {
			if	( ++LineNum > HEIGHT-2 )
				break;
			for	( I = 1  ; ( I < MAX_COL ) && ( I <= Reader.FieldCount )  ; I++ ) {
				Sizes[ I ]		=	( Reader[ I ].Trim().Length > Sizes[ I ] )
							?	Reader[ I ].Trim().Length
							:	Sizes[ I ] ;
				Lines[ LineNum - 1 ]	+= (
								( I > 1 )
							?	TAB.ToString()
							:	CAbc.EMPTY
							)
							+	__.FixUkrI( Reader[ I ].Trim() );
			}
		}
		Reader.Close();
		if	( LineNum==0 )
			return	false;
		string	Line;
		string[]Columns;
		CConsole.Clear();
		TotalColumns	=	0;
		for	( I = 0  ; I < LineNum ;  I ++ ) {
			if	( Lines[ I ] == null )
				break;
			if	( Lines[ I ] == CAbc.EMPTY )
				break;
			Columns	=	Lines[ I ].Split( TAB );
			if	( Columns == null )
				continue;
			if	( Columns.Length == 0 )
				continue;
			Line	=	" ";
			int	J=0;
			foreach( string Column in Columns ) {
				Line	+= (
						( J > 0 )
					?	" | "
					:	CAbc.EMPTY
					)
					+	__.Left( Column , Sizes[ J+1 ] ) ;
				J++;
			}
			TotalColumns	=	(J>TotalColumns) ? J : TotalColumns;
			__.Print( __.Left( Line , WIDTH - 1 ) );
		}
		if	( TotalColumns==0 )
			return false;
		__.Print( __.Replicate("_",WIDTH-1),"Для продолжения нажмите ENTER.Для выхода - ESC.");
		if	( ! CConsole.WaitForEscOrEnter() )
			return	false;
		for	( int J=0 ; J<TotalColumns ; J++ ) {
			CConsole.Clear();
			for	( I = 0  ; ( I < LineNum ) && ( I < HEIGHT - 1 ) ;  I ++ ) {
				Columns	=	Lines[ I ].Split( TAB );
				if	( Columns == null )
					continue;
				if	( Columns.Length <= J )
					continue;
				__.Print( __.Left( Columns[ J ] , WIDTH - 1 ) );
			}
			int	MenuCount	=	1;
			if	( FieldFromFile[ FLD_SUMA         ] ==  0 )
				MenuCount ++ ;
			if	( FieldFromFile[ FLD_CREDITACC    ] ==	0 )
				MenuCount ++ ;
			if	( FieldFromFile[ FLD_CREDITNAME   ] ==  0 )
				MenuCount ++ ;
			if	( FieldFromFile[ FLD_CREDITSTATE  ] ==  0 )
				MenuCount ++ ;
			string[]MenuItems	= new	string[ MenuCount ];
			int[]	MenuKinds	= new	int[ MenuCount ];
			I			=	0;
			MenuItems[I]		=	" ( пропустить ) ";
			MenuKinds[I]		=	0 ;
			if	( FieldFromFile[ FLD_SUMA ] == 0 ) {
				I ++ ;
				MenuItems[ I ]	=	"     сумма" ;
				MenuKinds[ I ]	=	FLD_SUMA ;
			}
			if	( FieldFromFile[ FLD_CREDITACC ] == 0 ) {
				I ++ ;
				MenuItems[ I ]	=	"   кредит-счет" ;
				MenuKinds[ I ]	=	FLD_CREDITACC ;
			}
			if	( FieldFromFile[ FLD_CREDITNAME ] == 0 ) {
				I ++ ;
				MenuItems[ I ]	=	" название кт.счета" ;
				MenuKinds[ I ]	=	FLD_CREDITNAME ;
			}
			if	( FieldFromFile[ FLD_CREDITSTATE ] == 0 ) {
				I ++ ;
				MenuItems[ I ]	=	"идент код. кт.счета" ;
				MenuKinds[ I ]	=	FLD_CREDITSTATE ;
			}
			Choice	=	CConsole.GetMenuChoice( MenuItems )	;
			if	( Choice == 0 )	 {
				break;
			}
			if	( Choice > 1 )
				FieldFromFile[ MenuKinds[ Choice - 1 ] ] = J + 1 ;
		}
		return	true;
	}//FOLD00

	//  Получение информации о дб.счете
	public	bool	GetDebitInfo( CCommand Command ) {//fold00
		string[]MenuItems	= new	string[6] ;
		bool	Choice		=	false ;
		string	CmdRes		=	CAbc.EMPTY
		,	CmdText		=	CAbc.EMPTY
		,	SavedCode	=	CAbc.EMPTY;
		do {
			do {
				if	( FieldValues[ FLD_DEBITACC ] == CAbc.EMPTY ) {
					__.Print( CAbc.EMPTY );
					__.Write( "Введите номер дебет.счета : " );
					FieldValues[ FLD_DEBITACC ]	=	__.Input().Trim();
				}
				CmdText		=	" select  Convert(Char(16),LTrim( a.StateCode ) ) + LTrim( a.ShortName ) "
						+	" from    dbo.SV_Accounts as a with ( NoLock ) where a.SubCount=0 "
						+	" and     a.Code='" + __.GetCodeByMoniker( FieldValues[ FLD_DEBITACC ] ) + "'" ;
				CmdRes		=	(string) Command.GetScalar( CmdText );
				if	( CmdRes == null )
					FieldValues[ FLD_DEBITACC ] = CAbc.EMPTY ;
				else
					if	( CmdText.Length < 16 )
						FieldValues[ FLD_DEBITACC ] = CAbc.EMPTY ;
					else	{
						FieldValues[ FLD_DEBITSTATE ]	=	CmdRes.Substring(0,16).Trim();
						FieldValues[ FLD_DEBITNAME ]	=	__.FixUkrI( CmdRes.Substring(16).Trim() );
					}
			} while	( FieldValues[ FLD_DEBITACC ] == CAbc.EMPTY ) ;
			MenuItems[0]	=	"Дебет.счет : " + FieldValues[ FLD_DEBITACC ] ;
			MenuItems[1]	=	"Идент.код  : " + FieldValues[ FLD_DEBITSTATE ] ;
			MenuItems[2]	=	__.Left( "'"+ FieldValues[ FLD_DEBITNAME ].Trim() +"'" , 36 ) ;
			MenuItems[3]	=	"_____________________________________";
			MenuItems[4]	=	"Для продолжения нажмите ENTER." ;
			MenuItems[5]	=	"Для отмены ESC.";
			Choice		=	CConsole.GetBoxChoice( MenuItems ) ;
			if	( ! Choice )
				FieldValues[ FLD_DEBITACC ]	=	CAbc.EMPTY;
			else
				if	( __.FileExists( ModelFileName ) )
					__.DeleteFile( ModelFileName );
		} while	( ! Choice ) ;
		CConsole.Clear();
		if	( ( FieldFromFile[ FLD_SUMA ] < 1 ) || ( FieldFromFile[ FLD_SUMA ] == null ) ) {
			byte	SavedColor		=	CConsole.BoxColor;
			CConsole.BoxColor	=	CConsole.RED*16 + CConsole.WHITE	;
			CConsole.GetBoxChoice( "Не указана колонка с суммой !" , "" ,"Для выхода нажмите ENTER.") ;
			CConsole.BoxColor	=	SavedColor	;
			return	false;
		}
		__.Write( " Назначение платежа ( " + FieldValues[ FLD_PURPOSE ] + ") : " );
		string	Purpose	=	__.Input().Trim();
		if	( Purpose.Length > 1 )
			FieldValues[ FLD_PURPOSE ] = __.FixUkrI( Purpose );
		__.Write( " Нумеровать документы с номера ( " + FieldValues[ FLD_CODE ] + " ) : " );
		string	Code	=	__.Input().Trim();
		if	( Code.Length > 0 )
			FieldValues[ FLD_CODE ] = Code;
		FieldValues[ FLD_CODE ]	=	( __.CLng( FieldValues[ FLD_CODE ] ) - 1 ).ToString();
		__.Write( " Вводить датой ( " + __.StrD( DayDate , 8 , 8 ).Replace(",",".") + " ) : " );
		Code		=	__.Input().Trim();
		if	( Code.Length > 0 )
			if	( __.GetDate( Code ) > 40000 )
				 DayDate	=	__.GetDate( Code ) ;
		//  -------------------------------------------------------------
		if	( ! Reader.Open( FileName , CharSet ) ) {
			Reader.Close();
			return	false;
		}
		TotalLines=0;
		SavedCode	=	FieldValues[ FLD_CODE ];
		while	( Reader.Read() ) {
			PrepareFields();
			TotalLines	++;
			TotalSum	+=	__.CLng( "0" + FieldValues[ FLD_SUMA ].Trim() );
		}
		FieldValues[ FLD_CODE ]	=	SavedCode;
		Reader.Close();
		bool	Result		=	CConsole.GetBoxChoice(
						" Всего строк : " + __.Right( TotalLines.ToString() , 11 )
					,	" Общая сумма : " + __.StrN( TotalSum / 100 , 11 ).Replace(",",".")
					,	"_________________________________"
					,	" Для продолжения нажмите ENTER."
					,	" Для выхода - ESC. "
					);
		return	Result;
	}//FOLD00

	//  Спросить у пользователя, надо ли применять данный шаблон
	void	ShowModel() {//fold00
		if	( ! __.FileExists( FileName ) )
			return;
		if	( __.FileExists( ModelFileName ) )
			__.DeleteFile( ModelFileName );
		ModelFile      = new   CCfgFile( FileName ) ;
		string	DebitMoniker    =       (string) ModelFile["AccountA_Text"];
		if	( DebitMoniker == null )
			DebitMoniker	=	CAbc.EMPTY;
		string	DebitName    =       (string) ModelFile["AName_Text"];
		if	( DebitName == null )
			DebitName	=	CAbc.EMPTY;
		string	Purpose    =       (string) ModelFile["Argument_Text"];
		if	( Purpose == null )
			Purpose		=	CAbc.EMPTY;
		CConsole.Clear();
		__.Print( CAbc.EMPTY
		,	"    Информация из шаблона :"
		,	"    -----------------------"
		,	" Дебет-счет           :  " + DebitMoniker
		,	" Название дебет-счета :  " + DebitName
		,	" Назначение платежа   :  " + Purpose
		);
		if	( CConsole.GetBoxChoice(	"Использовать теперь этот шаблон ?"
						,	" Да = Enter . Нет = Esc ."
						)
			)
			__.CopyFile( FileName , ModelFileName );
	}//FOLD00

	//  Копируем промжуточный CSV полученный из DIF
	public	void	CopyTempFile( string TmpFileName ) {//fold00
		if	( NeedDelSrcFile )
			__.CopyFile( FileName , TmpFileName + ".CSV" );
	}//FOLD00

	//  Преобразование DIF  в CSV
	bool	ConvertDifToCsv() {//fold00
		string		TmpFileName	=	__.GetTempName() /* + ".CSV" */ ;
		int		Char_Set	=	CAbc.CHARSET_DOS;
		string		DELIMITER	=	";"	;
		string		QUOTE		=	CAbc.QUOTE;
		string		DOUBLEQUOTE	=	CAbc.QUOTE + CAbc.QUOTE;
		string		Value		=	""	;
                bool		WaitForAValue	=	false	;
                bool		HasDataStarted	=	false	;
                CTextReader	TextReader	= new	CTextReader();
                CTextWriter	TextWriter	= new	CTextWriter();
		if	( ! TextWriter.Create( TmpFileName , CAbc.CHARSET_WINDOWS ) )
			return	false;
		if	( TextReader.Open( FileName , Char_Set ) )
			while	( TextReader.Read() ) {
				Value	=	TextReader.Value.Trim();
                                if	( Value.Length < 3 )
                                	continue;
				if	( CCommon.Upper( Value )=="EOD" )
					break;
				if	( CCommon.Upper( Value )=="BOT" )
                                	if	( HasDataStarted )
						TextWriter.Add(  CAbc.CRLF );
                                        else	HasDataStarted	=	true;
                                if	( ! HasDataStarted )
                                	continue;
				if	( Value.Substring(0,2)=="0," )
					TextWriter.Add( Value.Substring(2).Replace(",",".") + DELIMITER );
				if	( Value.Substring(0,2)=="1," )
					WaitForAValue = true ;
				if	( WaitForAValue == true ) {
					if	( Value.Substring(0,1)== QUOTE )
						TextWriter.Add( Value.Substring( 1 , Value.Length-2 ).Replace( QUOTE + QUOTE , QUOTE ) + DELIMITER );
                                        WaitForAValue = true;
				}
			}
		else	{
			TextWriter.Close();
			return	false;
		}
		TextReader.Close();
		TextWriter.Close();
		FileName=TmpFileName;
		NeedDelSrcFile=true;
		return	true;
	}//FOLD00
}