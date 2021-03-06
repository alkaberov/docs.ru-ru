---
title: "Время запуска приложения | Microsoft Docs"
ms.custom: ""
ms.date: "03/30/2017"
ms.prod: ".net-framework"
ms.reviewer: ""
ms.suite: ""
ms.technology: 
  - "dotnet-wpf"
ms.tgt_pltfrm: ""
ms.topic: "article"
helpviewer_keywords: 
  - "запуск приложения [WPF]"
  - "производительность [WPF], время запуска"
  - "заставка [WPF], время запуска"
  - "время запуска [WPF]"
  - "WPF, время запуска"
ms.assetid: f0ec58d8-626f-4d8a-9873-c20f95e08b96
caps.latest.revision: 11
author: "dotnet-bot"
ms.author: "dotnetcontent"
manager: "wpickett"
caps.handback.revision: 11
---
# Время запуска приложения
Время, необходимое для запуска WPF\-приложения, может меняться в широких пределах.  В данном разделе описаны различные методы для сокращения воспринимаемого и фактического времени запуска WPF\-приложения.  
  
## Общие сведения о "холодном" и "горячем" запуске  
 "Холодный" запуск выполняется при первом запуске приложения после перезагрузки системы либо при запуске приложения, завершении его работы и повторном запуске по прошествии продолжительного времени.  Если при запуске приложения необходимые ему страницы \(код, статические данные, реестр и т. д.\) отсутствуют в списке простаивающих страниц диспетчера памяти Windows, возникают ошибки страницы.  Для загрузки таких страниц в память требуется обращение к жесткому диску.  
  
 "Горячий" запуск выполняется, когда большинство страниц основных компонентов среды CLR уже загружены в память, что позволяет не обращаться к жесткому диску и сэкономить таким образом время.  Именно поэтому управляемое приложение во второй раз запускается быстрее.  
  
## Реализация экрана\-заставки  
 В тех случаях, когда значительной задержки между запуском приложения и отображением графического интерфейса избежать невозможно, можно оптимизировать воспринимаемое время запуска с помощью *экрана\-заставки*.  При использовании такого подхода изображение выводится на экран практически сразу же после запуска пользователем приложения.  Когда приложение готово к выводу графического интерфейса, экран\-заставка исчезает.  Начиная с версии [!INCLUDE[net_v35SP1_short](../../../../includes/net-v35sp1-short-md.md)], класс <xref:System.Windows.SplashScreen> можно использовать для реализации экрана\-заставки.  Дополнительные сведения см. в разделе [Добавление в WPF\-приложение экрана\-заставки](../../../../docs/framework/wpf/app-development/how-to-add-a-splash-screen-to-a-wpf-application.md).  
  
 Собственный экран\-заставу также можно реализовать с помощью собственной графики Win32.  Собственную реализацию необходимо вывести на экран перед вызовом метода <xref:System.Windows.Application.Run%2A>.  
  
## Анализ кода запуска  
 Определите причину медленного "холодного" запуска.  Одной из причин \(но не единственной причиной\) медленного запуска могут быть операции дискового ввода\-вывода.  В общем случае следует максимально сократить использование внешних ресурсов, например сети, веб\-служб и жесткого диска.  
  
 Перед тестированием убедитесь в том, что никакие другие приложения или службы не используют управляемый код или WPF\-код.  
  
 Запустите WPF\-приложение сразу же после перезагрузки и засеките время, необходимое для запуска.  Если при последующих запусках \("горячий" запуск\) приложение загружается намного быстрее, скорее всего, причиной долгого "холодного" запуска являются операции дискового ввода\-вывода.  
  
 Если при "холодном" запуске приложения не используются операции дискового ввода\-вывода, вероятно, приложение долго инициализирует данные, выполняет расчеты, ожидает завершение какого\-либо события или выполняет при запуске JIT\-компиляцию.  В следующих разделах эти ситуации рассмотрены более подробно.  
  
## Оптимизация загрузки модулей  
 Чтобы определить загружаемые приложением модули, воспользуйтесь такими средствами как \(Procexp.exe\) и Tlist.exe.  Команда `Tlist <pid>` выводит все модули, загруженные указанным процессом.  
  
 Например, если компьютер не подключен к Интернету, а приложение загружает библиотеку System.Web.dll, значит, в приложении имеется модуль, который ссылается на эту сборку.  Убедитесь в том, что такая ссылка необходима.  
  
 Если в приложении используется несколько модулей, объедините их в один модуль.  При использовании такого подхода снижаются накладные расходы на загрузку сборок среды CLR.  Меньшее количество сборок также позволяет снизить затраты на поддержку их состояния средой CLR.  
  
## Приостановка операций инициализации  
 Выполните код инициализации после отображения основного окна приложения.  
  
 Имейте в виду, что инициализация может выполнять в конструкторе класса, а если код инициализации ссылается на другие классы, это может привести к каскадному эффекту, при котором будет запущено множество конструкторов.  
  
## Отказ от настройки приложения  
 Следует избегать настройки приложения.  Например, если у приложения простые требования к конфигурации и жесткие требования к времени запуска, для ускорения запуска можно использовать записи в реестре или простой INI\-файл.  
  
## Использование глобального кэша сборок  
 Если сборка не установлена в глобальном кэше сборок, это может привести к задержкам, связанным с проверкой хэша сборок со строгими именами и проверкой образа Ngen, если на компьютере доступен образ этой сборки в машинном коде.  Для сборок, установленных в глобальном кэше сборок, проверка строгих имен не выполняется.  Дополнительные сведения см. в разделе [Gacutil.exe \(программа глобального кэша сборок\)](../../../../docs/framework/tools/gacutil-exe-gac-tool.md).  
  
## Использование Ngen.exe  
 Используйте в приложении генератор образов в машинном коде \(Ngen.exe\).  Использование средства Ngen.exe позволяет снизить загрузку ЦП за счет дополнительного места на диске, поскольку образы в машинном коде, создаваемые средством Ngen.exe, как правило, имеют больший размер, чем образы MSIL.  
  
 Чтобы сократить время "горячего" запуска, всегда следует использовать в приложении средство Ngen.exe, поскольку при этом время ЦП не тратится на JIT\-компиляцию кода приложения.  
  
 В некоторых случаях использование средства Ngen.exe также полезно и для ускорения "холодного" запуска.  Это связано с тем, что при использовании данного средства JIT\-компилятор \(mscorjit.dll\) не загружается.  
  
 Одновременное использование модулей Ngen и JIT может привести к противоположному эффекту.  Это связано с необходимостью загрузки библиотеки mscorjit.dll, а при обработке кода JIT\-компилятором в процессе считывания метаданных сборок необходим доступ ко многим страницам, хранящимся в образах Ngen.  
  
### Технологии Ngen и ClickOnce  
 От способа развертывания приложения также может зависеть время его загрузки.  Развертывание приложения с помощью [!INCLUDE[ndptecclick](../../../../includes/ndptecclick-md.md)] не поддерживает Ngen.  Если в приложении будет использоваться технология Ngen.exe, следует выбрать другой механизм развертывания, например установщик Windows.  
  
 Дополнительные сведения см. в разделе [Ngen.exe \(генератор образов в машинном коде\)](../../../../docs/framework/tools/ngen-exe-native-image-generator.md).  
  
### Изменение базового адреса и конфликты адресов библиотек DLL  
 При использовании средства Ngen.exe следует иметь в виду, что при загрузке в память образов в машинном коде может быть изменен базовый адрес.  Если библиотека DLL не загружена по основному базовому адресу из\-за того, что диапазон адресов уже выделен, загрузчик Windows загрузит эту библиотеку по другому адресу, что может занять продолжительное время.  
  
 Чтобы проверить наличие модулей, в которых все страницы являются закрытыми, можно воспользоваться средством Virtual Address Dump \(Vadump.exe\).  Если такие модули существуют, возможно, они были загружены по другому базовому адресу.  Следовательно, страницы таких модулей нельзя совместно использовать.  
  
 Дополнительные сведения о задании базового адреса см. в разделе [Ngen.exe \(генератор образов в машинном коде\)](../../../../docs/framework/tools/ngen-exe-native-image-generator.md).  
  
## Оптимизация Authenticode  
 Проверка Authenticode увеличивает время запуска.  Сборки, подписанные с использованием технологии Authenticode, в обязательном порядке проверяются центром сертификации.  Такая проверка может занять продолжительное время, поскольку для ее выполнения может потребоваться несколько раз подключиться к сети, чтобы загрузить текущие списки отзыва сертификатов.  При этом также проверяется наличие полной цепочки допустимых сертификатов в пути к доверенному корневому центру сертификации.  Выполнение такой проверки при загрузке сборки может занять несколько секунд.  
  
 Установите на клиентский компьютер сертификат центра сертификации или откажитесь от использования Authenticode, если это возможно.  Если приложению не требуется свидетельство издателя, не нужно проверять подпись, расплачиваясь за это увеличением времени запуска.  
  
 Начиная с версии [!INCLUDE[net_v35_short](../../../../includes/net-v35-short-md.md)], можно использовать параметр конфигурации, позволяющий отключить проверку Authenticode.  Для этого добавьте в файл app.exe.config следующий параметр:  
  
```  
<configuration>  
    <runtime>  
        <generatePublisherEvidence enabled="false"/>   
    </runtime>  
</configuration>  
```  
  
 Дополнительные сведения см. в разделе [Элемент \<generatePublisherEvidence\>](../../../../docs/framework/configure-apps/file-schema/runtime/generatepublisherevidence-element.md).  
  
## Сравнение производительности с Windows Vista  
 В диспетчере памяти Windows Vista используется технология SuperFetch.  Технология SuperFetch анализирует шаблоны использования памяти за длительное время и определяет оптимальное содержимое памяти для конкретного пользователя.  Эта технология работает непрерывно и постоянно поддерживает память в оптимальном состоянии.  
  
 Данный подход отличается от используемой в Windows XP технологии предварительной выборки, при которой данные в память загружались без анализа шаблонов использования.  Со временем, если пользователь часто работает с WPF\-приложением в Windows Vista, время "холодного" запуска приложения может сократиться.  
  
## Эффективное использование доменов приложения  
 Если возможно, загружайте сборки в нейтральную по отношению к доменам область кода, чтобы образ в машинном коде \(если таковой имеется\) гарантированно использовался во всех созданных в приложении доменах приложения.  
  
 Для повышения производительности обеспечьте эффективное взаимодействие между доменами, сократив количество вызовов между домена.  Если возможно, используйте вызовы без аргументов либо с аргументами примитивных типов.  
  
## Использование атрибута NeutralResourcesLanguage  
 Используйте атрибут <xref:System.Resources.NeutralResourcesLanguageAttribute>, чтобы указать для диспетчера <xref:System.Resources.ResourceManager> нейтральные язык и региональные параметры.  Такой подход позволяет избежать неудачного поиска сборок.  
  
## Использование для сериализации класса BinaryFormatter  
 Если необходимо использовать сериализацию, воспользуйтесь вместо класса <xref:System.Xml.Serialization.XmlSerializer> классом <xref:System.Runtime.Serialization.Formatters.Binary.BinaryFormatter>.  Класс <xref:System.Runtime.Serialization.Formatters.Binary.BinaryFormatter> реализован в библиотеке базовых классов \(BCL\) сборки mscorlib.dll.  Класс <xref:System.Xml.Serialization.XmlSerializer> реализован в сборке System.Xml.dll, что может привести к загрузке дополнительной библиотеки DLL.  
  
 Если необходимо использовать класс <xref:System.Xml.Serialization.XmlSerializer>, то повысить производительность можно путем предварительного создания сборки сериализации.  
  
## Настройка технологии ClickOnce на проверку обновлений после запуска  
 Если в приложении используется технология [!INCLUDE[ndptecclick](../../../../includes/ndptecclick-md.md)], не подключайтесь к сети при запуске приложения, настроив [!INCLUDE[ndptecclick](../../../../includes/ndptecclick-md.md)] на веб\-узла развертывания на наличие обновлений после запуска приложения.  
  
 При использовании модели приложения\-браузера XAML \(XBAP\), имейте в виду, что технология [!INCLUDE[ndptecclick](../../../../includes/ndptecclick-md.md)] проверяет веб\-узел развертывания на наличие обновлений даже в том случае, если XBAP уже находится в кэше [!INCLUDE[ndptecclick](../../../../includes/ndptecclick-md.md)].  Дополнительные сведения см. в разделе [Развертывание и безопасность технологии ClickOnce](../Topic/ClickOnce%20Security%20and%20Deployment.md).  
  
## Настройка автоматического запуска службы PresentationFontCache  
 Первым WPF\-приложением, которое следует запустить после перезагрузки, является служба PresentationFontCache.  Эта служба кэширует системные шрифты, ускоряет доступ к шрифтам и повышает общую производительность.  С запуском службы связаны дополнительные накладные расходы, и в некоторых управляемых средах следует настроить эту службу на автоматический запуск после перезагрузки системы.  
  
## Использование программной привязки данных  
 Вместо использования XAML для установки контекста <xref:System.Windows.FrameworkElement.DataContext%2A> для главного окна приложения декларативным образом задайте этот контекст программно в методе <xref:System.Windows.Application.OnActivated%2A>.  
  
## См. также  
 <xref:System.Windows.SplashScreen>   
 <xref:System.AppDomain>   
 <xref:System.Resources.NeutralResourcesLanguageAttribute>   
 <xref:System.Resources.ResourceManager>   
 [Добавление в WPF\-приложение экрана\-заставки](../../../../docs/framework/wpf/app-development/how-to-add-a-splash-screen-to-a-wpf-application.md)   
 [Ngen.exe \(генератор образов в машинном коде\)](../../../../docs/framework/tools/ngen-exe-native-image-generator.md)   
 [Элемент \<generatePublisherEvidence\>](../../../../docs/framework/configure-apps/file-schema/runtime/generatepublisherevidence-element.md)