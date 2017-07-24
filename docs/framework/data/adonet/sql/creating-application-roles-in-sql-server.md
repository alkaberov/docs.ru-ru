---
title: "Создание ролей приложения в SQL Server | Microsoft Docs"
ms.custom: ""
ms.date: "03/30/2017"
ms.prod: ".net-framework-4.6"
ms.reviewer: ""
ms.suite: ""
ms.technology: 
  - "dotnet-ado"
ms.tgt_pltfrm: ""
ms.topic: "article"
ms.assetid: 27442435-dfb2-4062-8c59-e2960833a638
caps.latest.revision: 9
author: "JennieHubbard"
ms.author: "jhubbard"
manager: "jhubbard"
caps.handback.revision: 9
---
# Создание ролей приложения в SQL Server
Роли приложения предоставляют возможность назначать разрешения приложению, а не ролям базы данных или пользователям.  Пользователи могут соединиться с базой данных, активировать роль приложения и воспользоваться разрешениями, предоставленными приложению.  Разрешения, предоставленные роли приложения, действуют в течение всего соединения.  
  
> [!IMPORTANT]
>  Роли приложения активируются, когда клиентское приложение передает имя роли приложения и пароль в строке соединения.  В двухуровневом приложении безопасность при этом снижается, поскольку пароль должен храниться на клиентском компьютере.  В трехуровневом приложении пароль можно хранить так, чтобы он не был доступен пользователям приложения.  
  
## Функции ролей приложения  
 Роли приложения имеют следующие функции.  
  
-   В отличие от ролей базы данных роли приложения не содержат элементов.  
  
-   Роли приложения активируются, когда клиентское приложение передает имя и пароль роли приложения системной хранимой процедуре `sp_setapprole`.  
  
-   Пароль должен храниться на клиентском компьютере и передаваться во время выполнения. Роль приложения нельзя активировать из SQL Server.  
  
-   Пароль не зашифрован.  При сохранении пароля используется одностороннее хэширование.  
  
-   После активации разрешения, полученные через роль приложения, действуют в течение всего соединения.  
  
-   Роль приложения наследует разрешения, предоставленные роли `public`.  
  
-   Если роль приложения активирует член предопределенной роли сервера `sysadmin`, контекст безопасности переключается на контекст безопасности роли приложения на все время соединения.  
  
-   Если создать учетную запись `guest` в базе данных, в которой имеется роль приложения, то не нужно будет создавать учетную запись для роли приложения или для любого имени входа, вызывающего эту роль.  Роли приложения могут получать прямой доступ к другой базе данных, только если во второй базе данных существует учетная запись `guest`.  
  
-   Встроенные функции, возвращающие имена входа, например SYSTEM\_USER, возвращают имя входа, которое вызвало роль приложения.  Встроенные функции, возвращающие имена пользователей базы данных, возвращают имя роли приложения.  
  
### Принцип минимальных привилегий  
 Ролям приложений следует предоставлять только необходимые разрешения на случай, если пароль будет раскрыт.  Разрешения роли `public` необходимо отменить во всех базах данных, использующих роль приложения.  Отключите учетную запись `guest` во всех базах данных, к которым не следует предоставлять доступ участникам роли приложения.  
  
### Расширения роли приложения  
 Контекст выполнения после активации роли приложения можно переключить обратно на исходного участника, устранив тем самым необходимость отключения пула соединений.  Хранимая процедура `sp_setapprole` получила новый параметр, создающий файл cookie, который содержит контекстные сведения об участнике.  Сеанс можно восстановить, вызвав процедуру `sp_unsetapprole` и передав ей этот файл cookie.  
  
## Альтернативы ролям приложений  
 Роли приложений зависят от безопасности пароля, который представляет собой потенциально уязвимое место системы безопасности.  Пароли уязвимы в том случае, если они внедрены в код приложения или хранятся на диске.  
  
 Можно рассмотреть следующие альтернативные подходы.  
  
-   Используйте переключение контекста с помощью инструкции EXECUTE AS с предложениями NO REVERT и WITH COOKIE.  В базе данных можно создать пользовательскую учетную запись, которая не сопоставлена с именем входа.  Этой учетной записи присваиваются разрешения.  Использование инструкции EXECUTE AS для пользователя, не имеющего имени входа, безопаснее, поскольку этот подход основан на разрешениях, а не на пароле.  Для получения дополнительной информации см. [Настройка разрешений в SQL Server с олицетворением](../../../../../docs/framework/data/adonet/sql/customizing-permissions-with-impersonation-in-sql-server.md).  
  
-   Подписывайте хранимые процедуры сертификатами, предоставляя разрешение только на выполнение этих процедур.  Для получения дополнительной информации см. [Подписывание хранимых процедур в SQL Server](../../../../../docs/framework/data/adonet/sql/signing-stored-procedures-in-sql-server.md).  
  
## Внешние ресурсы  
 Дополнительные сведения см. в следующих ресурсах.  
  
|Ресурс|Описание|  
|------------|--------------|  
|[Роли приложений](http://msdn.microsoft.com/library/ms190998.aspx) в электронной документации по SQL Server|Описывает процесс создания и использования ролей приложения в SQL Server 2008.|  
  
## См. также  
 [Защита приложений ADO.NET](../../../../../docs/framework/data/adonet/securing-ado-net-applications.md)   
 [Общие сведения о безопасности в SQL Server](../../../../../docs/framework/data/adonet/sql/overview-of-sql-server-security.md)   
 [Сценарии защиты приложений в SQL Server](../../../../../docs/framework/data/adonet/sql/application-security-scenarios-in-sql-server.md)   
 [Центр разработчиков, поставщики ADO.NET Managed Provider и набор данных](http://go.microsoft.com/fwlink/?LinkId=217917)