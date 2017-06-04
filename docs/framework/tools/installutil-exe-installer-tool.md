---
title: "Installutil.exe (Installer Tool) | Microsoft Docs"
ms.custom: ""
ms.date: "03/30/2017"
ms.prod: ".net-framework"
ms.reviewer: ""
ms.suite: ""
ms.technology: 
  - "dotnet-clr"
ms.tgt_pltfrm: ""
ms.topic: "article"
dev_langs: 
  - "VB"
  - "CSharp"
  - "C++"
  - "jsharp"
helpviewer_keywords: 
  - "uninstalling server resources"
  - "removing server resources"
  - "status information for installation"
  - "installation progress reports"
  - "installing server resources"
  - "Installer tool"
  - "Installutil.exe"
  - "files, Installer tool"
  - "progress information for installation"
  - "reporting installation progress"
ms.assetid: 3f9d0533-f895-4897-b4ea-528284e0241d
caps.latest.revision: 40
author: "rpetrusha"
ms.author: "ronpet"
manager: "wpickett"
caps.handback.revision: 38
---
# Installutil.exe (Installer Tool)
Программа для установки является программой командной строки, с помощью которой можно устанавливать и удалять ресурсы сервера путем выполнения компонентов установщика в соответствующих сборках.  Эта программа работает совместно с классами в пространстве имен <xref:System.Configuration.Install>.  
  
 Эта программа автоматически устанавливается вместе с Visual Studio.  Программу можно запустить из командной строки разработчика \(или из командной строки Visual Studio в Windows 7\).  Дополнительные сведения см. в разделе [Командные строки](../../../docs/framework/tools/developer-command-prompt-for-vs.md).  
  
 В командной строке введите следующее.  
  
## Синтаксис  
  
```  
  
installutil [/u[ninstall]] [options] assembly [[options] assembly] ...   
```  
  
#### Параметры  
  
|Аргумент|Описание|  
|--------------|--------------|  
|`assembly`|Имя файла сборки, в которой должны выполняться компоненты установщика.  Пропустите этот параметр, если указывается строгое имя сборки с помощью параметра `/AssemblyName`.|  
  
<a name="options"></a>   
## Параметры  
  
|Параметр|Описание|  
|--------------|--------------|  
|`/h[elp]`<br /><br /> – или –<br /><br /> `/?`|Отображает синтаксис команд и параметров программы.|  
|`/help` *assembly*<br /><br /> – или –<br /><br /> `/?` *assembly*|Отображает дополнительные параметры, распознаваемые отдельными установщиками в пределах указанной сборки, вместе с синтаксисом команд и параметров для программы InstallUtil.exe.  Этот параметр добавляет текст, возвращенный каждым свойством компонента установщика <xref:System.Configuration.Install.Installer.HelpText%2A?displayProperty=fullName>, в текст справки программы InstallUtil.exe.|  
|`/AssemblyName` "*assemblyName*<br /><br /> ,Version\=*major.minor.build.revision*<br /><br /> ,Culture\=*locale*<br /><br /> ,PublicKeyToken\=*publicKeyToken*"|Задает строгое имя сборки, которое требуется зарегистрировать в глобальном кэше сборок.  Имя сборки должно содержать версию, язык и региональные параметры, а также токен открытого ключа сборки.  Полное имя должно быть заключено в кавычки.<br /><br /> Например, "myAssembly, Culture\=neutral, PublicKeyToken\=0038abc9deabfle5, Version\=4.0.0.0" — это полное имя сборки.|  
|`/InstallStateDir=[` *directoryName* `]`|Задает каталог InstallState\-файла, содержащего данные, которые используются для удаления сборки.  По умолчанию используется каталог, содержащий сборку.|  
|`/LogFile=`\[*filename*\]|Задает имя файла журнала, в который записывается ход установки.  По умолчанию, если параметр `/LogFile` не указан, создается файл журнала с именем *assemblyname*.InstallLog.  Если параметр *filename* не указан, файл журнала не создается.|  
|`/LogToConsole`\={`true`&#124;`false`}|Если значение — `true`, выходные данные отображаются на консоли.  Если значение — `false` \(значение по умолчанию\), выходные данные на консоль не выводятся.|  
|`/ShowCallStack`|Если в ходе установки возникает исключение, содержимое стека вызовов заносится в файл журнала.|  
|`/u`\[`ninstall`\]|Удаляет указанные сборки.  В отличие от других параметров, `/u` применяется ко всем сборкам независимо от того, где этот параметр указан в командной строке.|  
  
<a name="cmdline"></a>   
## Дополнительные параметры установщика  
 Отдельные установщики, используемые в сборке, кроме перечисленных в разделе [Параметры](#options) могут распознавать и другие параметры.  Чтобы узнать об этих параметрах, запустите программу InstallUtil.exe с путями к сборкам в командной строке, а также с параметром `/?` или `/help`.  Чтобы задать эти параметры, необходимо включить их в командную строку вместе с параметрами, распознаваемыми программой InstallUtil.exe.  
  
> [!NOTE]
>  Текст справки о параметрах, поддерживаемых отдельными компонентами установщика, возвращается свойством <xref:System.Configuration.Install.Installer.HelpText%2A?displayProperty=fullName>.  Отдельные параметры, которые были введены в командной строке, доступны программно из свойства <xref:System.Configuration.Install.Installer.Context%2A?displayProperty=fullName>.  
  
 Все параметры командной строки записываются в файл журнала установки.  Однако если используется параметр `/Password`, распознаваемый некоторыми компонентами установщика, сведения о пароле будут заменены восемью звездочками \(\*\) и не будут отображаться в файле журнала.  
  
> [!IMPORTANT]
>  В некоторых случаях передаваемые в установщик параметры могут содержать конфиденциальные или личные сведения, которые по умолчанию записываются в обычный текстовый файл журнала.  Для предотвращения этого поведения можно запретить ведение журнала, указав в командной строке `/LogFile=` \(без аргумента *filename*\) после программы Installutil.exe.  
  
## Заметки  
 Приложения .NET Framework состоят из традиционных файлов программ и связанных с ними ресурсов, таких как очереди сообщений, журналы событий и счетчики производительности, которые создаются при развертывании приложения.  Компоненты установщика сборки могут использоваться для создания таких ресурсов при установке приложения и для их удаления при удалении приложения.  Программа Installutil.exe обнаруживает и выполняет эти компоненты установщика.  
  
 В командной строке можно указать сразу несколько сборок.  Параметры должны указываться перед именами сборок, к установке которых они относятся.  За исключением `/u` и `/AssemblyName`, параметры являются накопительными, но переопределяемыми.  То есть параметры, указанные для одной сборки, применяются и ко всем последующим сборкам. Исключением являются параметры, с которыми указывается новое значение.  
  
 При запуске программы Installutil.exe для сборки без указания параметров она помещает в каталог сборки следующие три файла.  
  
-   InstallUtil.InstallLog — содержит общее описание хода выполнения установки.  
  
-   *assemblyname*.InstallLog — содержит сведения, относящиеся к этапу фиксации процесса установки.  Дополнительные сведения об этапе фиксации см. в описании метода <xref:System.Configuration.Install.Installer.Commit%2A>.  
  
-   *assemblyname*.InstallState — содержит данные, используемые для удаления сборки.  
  
 Программа Installutil.exe использует отражение для проверки указанных сборок и поиска всех типов <xref:System.Configuration.Install.Installer>, у которых значение атрибута <xref:System.ComponentModel.RunInstallerAttribute?displayProperty=fullName> имеет значение `true`.  Программа, выполняющая метод <xref:System.Configuration.Install.Installer.Install%2A?displayProperty=fullName> или <xref:System.Configuration.Install.Installer.Uninstall%2A?displayProperty=fullName> для каждого экземпляра типа <xref:System.Configuration.Install.Installer>.  Программа Installutil.exe выполняет установку как транзакцию, то есть если какую\-либо сборку не удалось установить, отменяется установка всех остальных сборок.  Удаление не считается транзакцией.  
  
 Программа Installutil.exe не может устанавливать или удалять сборки с отложенной подписью, но может устанавливать и удалять сборки со строгими именами.  
  
 Начиная с .NET Framework версии 2.0, 32\-разрядная версия среды CLR поставляется только с 32\-разрядной версией программы установщика, однако 64\-разрядная версия среды CLR поставляется и с 32\-разрядной, и с 64\-разрядной версиями программы установщика.  При работе с 64\-разрядной средой CLR используйте 32\-разрядную программу установщика для установки 32\-разрядных сборок, а 64\-разрядную программу установщика — для установки 64\-разрядных СIL\-сборок.  Обе версии программы установщика ведут себя одинаково.  
  
 Программу Installutil.exe невозможно использовать для развертывания службы Windows, написанной на языке C\+\+, потому что программа Installutil.exe не может распознать встроенный машинный код, созданный компилятором C\+\+.  При попытке развернуть службу Windows, написанную на языке C\+\+, с помощью программы Installutil.exe возникает исключение, например <xref:System.BadImageFormatException>.  Для работы с этим сценарием перенесите код службы в модуль C\+\+ и создайте объект установщика на языке C\# или Visual Basic.  
  
## Примеры  
 Следующая команда выводит описание синтаксиса и параметров команды для программы InstallUtil.exe.  
  
```  
installutil /?  
```  
  
 Следующая команда выводит описание синтаксиса и параметров команды для программы InstallUtil.exe.  Она также отображает для `myAssembly.exe` описание и список параметров, поддерживаемых компонентами установщика, если свойству <xref:System.Configuration.Install.Installer.HelpText%2A?displayProperty=fullName> установщика было задано значение текста справки.  
  
```  
installutil /? myAssembly.exe  
```  
  
 Следующая команда выполняет компоненты установщика в сборке `myAssembly.exe`.  
  
```  
installutil myAssembly.exe  
```  
  
 Следующая команда выполняет компоненты установщика в сборке с помощью ключа `/AssemblyName` и полного имени.  
  
```  
installutil /AssemblyName "myAssembly, Culture=neutral, PublicKeyToken=0038abc9deabfle5, Version=4.0.0.0"  
```  
  
 Следующая команда выполняет компоненты установщика в сборке, заданной по имени файла, и в сборке, заданной по строгому имени.  Обратите внимание, что все сборки, указанные по имени файла, должны предшествовать сборкам, указанным по строгому имени в командной строке, потому что параметр `/AssemblyName` не может быть переопределен.  
  
```  
installutil myAssembly.exe /AssemblyName "myAssembly, Culture=neutral, PublicKeyToken=0038abc9deabfle5, Version=4.0.0.0"  
```  
  
 Следующая команда выполняет компоненты программы удаления в сборке `myAssembly.exe`.  
  
```  
installutil /u myAssembly.exe   
```  
  
 Следующая команда выполняет компоненты программы удаления в сборках `myAssembly1.exe` и `myAssembly2.exe`.  
  
```  
installutil myAssembly1.exe /u myAssembly2.exe  
```  
  
 Поскольку позиция параметра `/u` в командной строке не имеет значения, результат аналогичен выполнению следующей команды.  
  
```  
installutil /u myAssembly1.exe myAssembly2.exe  
```  
  
 Следующая команда выполняет компоненты установщика в сборке `myAssembly.exe` и указывает, что сведения о ходе установки должны записываться в файл `myLog.InstallLog`.  
  
```  
installutil /LogFile=myLog.InstallLog myAssembly.exe   
```  
  
 Следующая команда запускает установщики в сборке `myAssembly.exe`, указывает, что сведения о ходе выполнения должны записываться в файл `myLog.InstallLog`, и использует настраиваемый параметр установщика `/reg` для указания, что обновления должны вноситься в системный реестр.  
  
```  
installutil /LogFile=myLog.InstallLog /reg=true myAssembly.exe  
```  
  
 Следующая команда запускает установщики в сборке `myAssembly.exe`, использует пользовательский параметр установщика `/email` для указания адреса электронной почты пользователя и не ведет запись в файл журнала.  
  
```  
installutil /LogFile= /email=admin@mycompany.com myAssembly.exe  
```  
  
 Следующая команда записывает сведения о ходе установки для сборки `myAssembly.exe` в журнал `myLog.InstallLog`, а сведения для сборки `myTestAssembly.exe` — в журнал `myTestLog.InstallLog`.  
  
```  
installutil /LogFile=myLog.InstallLog myAssembly.exe /LogFile=myTestLog.InstallLog myTestAssembly.exe  
```  
  
## См. также  
 <xref:System.Configuration.Install>   
 [Tools](../../../docs/framework/tools/index.md)   
 [Командные строки](../../../docs/framework/tools/developer-command-prompt-for-vs.md)