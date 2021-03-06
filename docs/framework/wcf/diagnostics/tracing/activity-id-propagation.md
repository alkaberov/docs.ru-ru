---
title: "Распространение ИД действия | Microsoft Docs"
ms.custom: ""
ms.date: "03/30/2017"
ms.prod: ".net-framework-4.6"
ms.reviewer: ""
ms.suite: ""
ms.technology: 
  - "dotnet-clr"
ms.tgt_pltfrm: ""
ms.topic: "article"
ms.assetid: cd1c1ae5-cc8a-4f51-90c9-f7b476bcfe70
caps.latest.revision: 6
author: "Erikre"
ms.author: "erikre"
manager: "erikre"
caps.handback.revision: 6
---
# Распространение ИД действия
Распространение происходит при включении трассировки действий ServiceModel \(распространение ServiceModel\) или ее выключении \(распространение действий от пользователя к пользователю\).  
  
## Трассировка действий ServiceModel включена  
 Если трассировка действий ServiceModel включена, между действиями ProcessAction осуществляется распространение.  
  
### Server  
 Если атрибуту `propagateActivity` присвоено значение `true` на клиенте и сервере, идентификатор действия `ProcessAction` на сервере идентичен идентификатору в распространенном заголовке сообщения `ActivityId`.  
  
 Если в сообщении отсутствует заголовок `ActivityId` \(т.е. `propagateActivity`\=`false` на клиенте\) или если на сервере `propagateActivity`\=`false`, сервер создает новый идентификатор действия.  
  
### Клиент  
 Если клиент синхронно однопоточный, он игнорирует любые параметры `propagateActivity` на уровне клиента или сервера.  Вместо этого ответ обрабатывается действием запроса.  Если клиент асинхронный или синхронный многопоточный, в нем `propagateActivity`\=`true`, и имеется заголовок с идентификатором действия в сообщении, отправленном сервером. Клиент извлекает из сообщения идентификатор действия и передает его в действие Process Action, содержащее распространенный идентификатор действия.  Иначе клиент передает его из действия Process Action в новое действие Process Action.  Такая дополнительная передача в новое действие Process Action осуществляется для обеспечения согласованности.  В этом действии клиент извлекает идентификатор действия BeginCall из контекста локального потока, когда выделяется поток для обработки ответного сообщения.  После этого осуществляется передача в первоначальное действие Process Action.  
  
 Если клиент дуплексный, он действует в качестве сервера для получения сообщения.  
  
### Распространение в сообщениях об ошибках  
 Разницы в обработке допустимых сообщений и сообщений об ошибках нет.  Если `propagateActivity`\=`true`, идентификатор действия, добавляемый в заголовки сообщений об ошибках SOAP, идентичен внешнему действию.  
  
## Трассировка действий ServiceModel отключена  
 Если трассировка действий ServiceModel отключена, распространение осуществляется непосредственно между действиями пользовательского кода, минуя действия ServiceModel.  Идентификатор пользовательского действия также распространяется через заголовок идентификатора действия сообщения.  
  
 Если `propagateActivity`\=`true` и `ActivityTracing`\=`off` для прослушивателя трассировки ServiceModel \(независимо от того, включена ли трассировка в ServiceModel\), с клиентом или сервером происходит следующее.  
  
-   При запросе операции или отправке ответа идентификатор действия в TLS распространяется из пользовательского кода, пока не будет сформировано сообщение.  Заголовок идентификатора действия также вставляется в сообщение перед отправкой.  
  
-   При получении запроса или ответа идентификатор действия извлекается из заголовка сообщения сразу после создания объекта полученного сообщения.  Идентификатор действия в TLS распространяется сразу после исчезновения сообщения из области, пока не будет достигнут пользовательский код.  
  
 Такие действия гарантируют, что при включенном распространении пользовательские трассировки появятся в том же действии.  Однако гарантии в отношении трассировок ServiceModel отсутствуют.  Трассировки ServiceModel возникают в действии пользовательского кода, только если обработка таких трассировок выполняется в том же потоке, в котором было задано действие пользовательского кода.  
  
 Как правило, трассировки ServiceModel находятся в следующих местах.  
  
-   Если трассировка ServiceModel отключена, все выданные трассировки находятся в действиях пользователя.  Если включено распространение на сервере и клиенте, эти трассировки будут находиться в том же действии.  
  
-   Если включена трассировка ServiceModel, но ActivityTracing отключена, пользовательские трассировки находятся в том же действии, если для обеих сторон включено распространение.  Трассировки ServiceModel появляются в действии по умолчанию 0000, пока они не возникнут в том же потоке, где выполняется обработка пользовательского кода с первоначально заданным действием.  
  
-   Если включены трассировка ServiceModel и ActivityTracing, пользовательские трассировки находятся в пользовательских действиях, а трассировка ServiceModel \- в действиях, определенных ServiceModel.  Распространение осуществляется на уровне ServiceModel.