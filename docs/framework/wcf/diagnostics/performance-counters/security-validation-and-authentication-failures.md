---
title: "Сбои при проверке безопасности и проверке подлинности | Microsoft Docs"
ms.custom: ""
ms.date: "03/30/2017"
ms.prod: ".net-framework-4.6"
ms.reviewer: ""
ms.suite: ""
ms.technology: 
  - "dotnet-clr"
ms.tgt_pltfrm: ""
ms.topic: "article"
ms.assetid: 0d4e3666-dfc6-421c-baf8-9479c22f7050
caps.latest.revision: 8
author: "BrucePerlerMS"
ms.author: "bruceper"
manager: "mbaldwin"
caps.handback.revision: 8
---
# Сбои при проверке безопасности и проверке подлинности
Имя счетчика: Security Validation and Authentication Failures  
  
## Описание  
 Значение этого счетчика увеличивается всякий раз, когда сообщение отклоняется из\-за проблемы безопасности, не относящейся к счетчику "Security Calls Not Authorized".К таким проблемам относятся следующие.  
  
-   Невозможно прочесть в этом сообщении маркер клиента.  
  
-   Маркер клиента не прошел проверку подлинности \(например, неправильный пароль\).  
  
-   Сбой при проверке подписи \(например, сообщение было искажено\).  
  
-   Сообщение является копией предыдущего, что может свидетельствовать об атаке воспроизведения.  
  
-   Произошел сбой при расшифровке.  
  
-   В сообщении отсутствуют некоторые обязательные элементы \(например, отсутствует отметка времени или зашифрованный блок данных\).  
  
-   Во время подтверждения TLSNEGO\/SPNEGO произошли ошибки.