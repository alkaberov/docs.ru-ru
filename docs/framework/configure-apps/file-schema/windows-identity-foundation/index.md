---
title: "Схема конфигурации Windows Identity Foundation"
ms.custom: 
ms.date: 03/30/2017
ms.prod: .net-framework
ms.reviewer: 
ms.suite: 
ms.technology:
- dotnet-clr
ms.tgt_pltfrm: 
ms.topic: article
ms.assetid: 4d4f6d76-49a5-4bad-b345-097b2e2844e9
caps.latest.revision: 6
author: BrucePerlerMS
ms.author: bruceper
manager: mbaldwin
ms.translationtype: HT
ms.sourcegitcommit: 306c608dc7f97594ef6f72ae0f5aaba596c936e1
ms.openlocfilehash: b2ac7e97627eba85013e1effdc4f856f3df79089
ms.contentlocale: ru-ru
ms.lasthandoff: 09/05/2017

---
# <a name="windows-identity-foundation-configuration-schema"></a>Схема конфигурации Windows Identity Foundation
В представленных в этом разделе статьях приводятся сведения о схеме конфигурации Windows Identity Foundation (WIF). Настроить использование WIF в приложении можно также с помощью классов, предоставляемых платформой. Эти классы указываются в разделах, посвященных соответствующим элементам схемы. Ниже показана базовая структура тегов XML, предоставляемая схемой конфигурации WIF. Атрибуты не приводятся. Выделенные комментарии указывают на основные компоненты схемы.  
  
```xml  
<system.identityModel>  
    <!-- Service Configuration -->  
    <identityConfiguration>  
        <caches>  
            <sessionSecurityTokenCache />  
            <tokenReplayCache />  
        </caches>  
  
        <certificateValidation>  
            <certificateValidator />   
        </certificateValidation>  
  
        <claimsAuthenticationManager />  
  
        <claimsAuthorizationManager>  
            <optionalConfigurationElement>  
        </claimsAuthorizationManager>  
  
        <claimTypeRequired>  
            <claimType />   
        </claimTypeRequired>  
  
        <tokenReplayDetection />  
  
        <!-- Security Token Handler Collection Configuration -->  
        <securityTokenHandlers>  
            <add>  
                <!-- Can take an optional configuration element which can be one of  
                     the following or a custom element -->  
                <samlSecurityTokenHandlerRequirement>  
                    <nameClaimType>  
                    <roleClaimType>   
                </samlSecurityTokenHandlerRequirement>  
  
                <sessionSecurityTokenHandlerRequirement />  
                <x509SecurityTokenHandlerRequirement />  
                <userNameSecurityTokenHandlerRequirement />  
            </add>  
            <clear />  
            <remove />  
            <securityTokenHandlerConfiguration>  
                <audienceUris>  
                    <add>  
                    <clear>  
                    <remove>  
                </audienceUris>  
  
                <caches>  
                    <sessionSecurityTokenCache />  
                    <tokenReplayCache />  
                </caches>  
  
                <certificateValidation>  
                    <certificateValidator>   
                </certificateValidation>  
  
                <issuerNameRegistry>  
                    <!-- Can take an optional configuration element which can be   
                         the <trustedIssuers> element to configure a configuration-based  
                         issuer name registry or can be a custom element -->  
                    <trustedIssuers>  
                        <add>  
                        <clear>  
                        <remove>  
                    </trustedIssuers>  
                </issuerNameRegistry>  
  
                <issuerTokenResolver />  
                <serviceTokenResolver />  
                <tokenReplayDetection />  
            </securityTokenHandlerConfiguration>  
        </securityTokenHandlers>  
    </identityConfiguration>  
</system.identityModel>  
  
<system.identityModel.services>  
    <!-- Federation Authentication Configuration -->  
    <federatedAuthentication>  
        <cookieHandler>  
            <chunkedCookieHandler />  
            <customCookieHandler />  
        </cookieHandler>  
  
        <serviceCertificate>  
            <certificateReference>  
        </serviceCertificate>  
  
        <wsFederation />  
    </federatedAuthentication>  
</system.identityModel.services>  
```  
  
## <a name="in-this-section"></a>Содержание  
 [\<system.identityModel>](../../../../../docs/framework/configure-apps/file-schema/windows-identity-foundation/system-identitymodel.md) Предоставляет конфигурацию для включения параметров WIF в приложениях.  
  
 [\<system.identityModel.services>](../../../../../docs/framework/configure-apps/file-schema/windows-identity-foundation/system-identitymodel-services.md) Предоставляет конфигурацию пассивной федерации посредством WIF. Настраивает модуль проверки подлинности сеансов (SAM) и модуль федеративной проверки подлинности (WSFAM).  
  
## <a name="related-sections"></a>Связанные разделы  
 [Настройка, администрирование и управление](http://msdn.microsoft.com/en-us/1e03c389-de2c-4096-aaff-86b087e1bea0) Описывается настройка приложений и служб WIF, а также управление ими.

