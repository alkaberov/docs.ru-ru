---
title: "Role-Based Security | Microsoft Docs"
ms.custom: ""
ms.date: "03/30/2017"
ms.prod: ".net"
ms.reviewer: ""
ms.suite: ""
ms.technology: 
  - "dotnet-standard"
ms.tgt_pltfrm: ""
ms.topic: "article"
dev_langs: 
  - "VB"
  - "CSharp"
  - "C++"
  - "jsharp"
helpviewer_keywords: 
  - "role-based security, about role-based security"
  - "user authentication, principals"
  - "principals [.NET Framework]"
  - "security [.NET Framework], role-based"
  - "permissions [.NET Framework], principals"
  - "authentication [.NET Framework], principals"
  - "role-based security, principals"
ms.assetid: 578cc32b-5654-4d8b-9d8c-ebcbc5c75390
caps.latest.revision: 13
author: "mairaw"
ms.author: "mairaw"
manager: "wpickett"
caps.handback.revision: 13
---
# Role-Based Security
Роли часто используются в финансовых и деловых приложениях для применения политики.  Например, приложение может установить ограничение на размер транзакции в зависимости от того, является ли сделавший запрос пользователь участником определенной роли.  Служащие могут иметь доступ к проведению транзакций, размер которых меньше определенного порога, менеджер может иметь больший предел, а вице\-президент — еще больший или неограниченный.  Безопасность на основе ролей также можно использовать, когда приложению для завершения действия требуются многократные утверждения.  В качестве примера можно привести систему закупок, в которой каждый сотрудник может создать заявку на покупку, но преобразовать эту заявку в заказ на покупку для отправки поставщику может только агент по закупкам.  
  
 Безопасность на основе ролей в .NET Framework поддерживает авторизацию путем предоставления текущему потоку информации о субъекте, который создается из связанного удостоверения.  Удостоверение \(и субъект, который определяется с его помощью\) может быть основано на учетной записи Windows или представлять собой пользовательское удостоверение, не связанное с учетной записью Windows.  Приложения .NET Framework могут принимать решения об авторизации либо на основе удостоверения субъекта, либо на основе членства в роли, либо на основе того и другого.  Роль — это именованный набор субъектов, имеющих одинаковые привилегии в плане безопасности \(например, кассир или менеджер\).  Субъект может иметь одну или несколько ролей.  Поэтому приложения могут использовать членство в ролях для определения того, имеет ли субъект право на выполнение запрошенного действия.  
  
 Для обеспечения простоты использования и согласованности с системой управления доступом для кода механизм безопасности на основе ролей в .NET Framework предоставляет объекты <xref:System.Security.Permissions.PrincipalPermission?displayProperty=fullName>, которые позволяют среде CLR проводить авторизацию таким же способом, как и при проверке управления доступом для кода.  Класс <xref:System.Security.Permissions.PrincipalPermission> представляет удостоверение или роль, которой должен соответствовать субъект. Этот класс совместим с декларативными и принудительными проверками безопасности.  При необходимости можно также получать доступ к данным удостоверения субъекта напрямую и выполнять проверки роли и удостоверения в коде.  
  
 Платформа .NET Framework обеспечивает поддержку безопасности на основе ролей, которая является достаточно гибкой и расширяемой для того, чтобы удовлетворять требованиям широкого спектра приложений.  Можно взаимодействовать с существующей инфраструктурой аутентификации, например службами COM\+ 1.0, или создать пользовательскую систему аутентификации.  Безопасность на основе ролей особенно хорошо подходит для использования в веб\-приложениях ASP.NET, которые обрабатываются в основном на сервере.  Тем не менее безопасность на основе ролей в .NET Framework можно использовать как на стороне клиента, так и на стороне сервера.  
  
 Перед прочтением этого раздела необходимо усвоить материал из раздела [Key Security Concepts](../../../docs/standard/security/key-security-concepts.md).  
  
## См. также  
  
|Заголовок|Описание|  
|---------------|--------------|  
|[Principal and Identity Objects](../../../docs/standard/security/principal-and-identity-objects.md)|Описываются принципы настройки универсальных удостоверений и субъектов и удостоверений и субъектов Windows, а также управления ими.|  
|[Key Security Concepts](../../../docs/standard/security/key-security-concepts.md)|Основные понятия, с которыми необходимо ознакомиться перед использованием механизмов безопасности .NET Framework.|  
  
## Ссылки  
 <xref:System.Security.Permissions.PrincipalPermission?displayProperty=fullName>