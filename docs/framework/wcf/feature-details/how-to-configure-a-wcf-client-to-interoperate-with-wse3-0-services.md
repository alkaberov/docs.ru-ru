---
title: "Практическое руководство. Настройка клиента WCF для взаимодействия со службами WSE&amp;3;.0 | Microsoft Docs"
ms.custom: ""
ms.date: "03/30/2017"
ms.prod: ".net-framework-4.6"
ms.reviewer: ""
ms.suite: ""
ms.technology: 
  - "dotnet-clr"
ms.tgt_pltfrm: ""
ms.topic: "article"
ms.assetid: 3dadd7f1-d207-4ea5-a73b-3e8aa44407f8
caps.latest.revision: 6
author: "Erikre"
ms.author: "erikre"
manager: "erikre"
caps.handback.revision: 6
---
# Практическое руководство. Настройка клиента WCF для взаимодействия со службами WSE&amp;3;.0
Клиенты [!INCLUDE[indigo1](../../../../includes/indigo1-md.md)] на уровне линий связи совместимы со службами расширений веб-служб версии 3.0 для Microsoft .NET (WSE), если клиенты [!INCLUDE[indigo2](../../../../includes/indigo2-md.md)] настроены на использование спецификации WS-Addressing версии августа 2004 г.  
  
### <a name="to-configure-a-wcf-client-to-interoperate-with-a-wse-30-web-service"></a>Настройка клиента WCF для взаимодействия с веб-службой WSE 3.0  
  
1.  Запустите [ServiceModel Metadata Utility Tool (Svcutil.exe)](../../../../docs/framework/wcf/servicemodel-metadata-utility-tool-svcutil-exe.md) для создания [!INCLUDE[indigo2](../../../../includes/indigo2-md.md)] клиента для WSE 3.0 веб-службы.  
  
     Для веб-службы WSE создается класс клиента [!INCLUDE[indigo2](../../../../includes/indigo2-md.md)].  
  
     Подробные сведения о создании [!INCLUDE[indigo2](../../../../includes/indigo2-md.md)] клиента, в разделе [Практическое руководство: создание клиента](../../../../docs/framework/wcf/how-to-create-a-wcf-client.md).  
  
2.  Создайте класс, представляющий привязку, которая может обмениваться данными с веб-службами WSE 3.0.  
  
     Следующий класс является частью [взаимодействие с WSE](http://msdn.microsoft.com/ru-ru/f6816861-96a0-45f9-8736-8e4e82cd3a41) образца.  
  
    1.  Создайте класс, производный от <xref:System.ServiceModel.Channels.Binding> класса.  
  
         В следующем примере кода создается класс с именем `WseHttpBinding` , производный от <xref:System.ServiceModel.Channels.Binding> класса.  
  
         [!code-csharp[c_WCFClientToWSEService#1](../../../../samples/snippets/csharp/VS_Snippets_CFX/c_wcfclienttowseservice/cs/wsehttpbinding.cs#1)]
         [!code-vb[c_WCFClientToWSEService#1](../../../../samples/snippets/visualbasic/VS_Snippets_CFX/c_wcfclienttowseservice/vb/wsehttpbinding.vb#1)]  
  
    2.  Добавьте в класс свойства, задающие готовое к использованию утверждение WSE, требуются ли производные ключи, используются ли безопасные сеансы, требуется ли подтверждение подписей, и параметры защиты сообщений.  
  
         В следующем примере кода определяется `SecurityAssertion,``RequireDerivedKeys, EstablishSecurityContext, MessageProtectionOrder` свойства, задающие готовое к использованию утверждение WSE, требуются ли производные ключи, используются ли безопасные сеансы, требуются ли подтверждение подписей и параметры защиты сообщений соответственно.  
  
         [!code-csharp[c_WCFClientToWSEService#3](../../../../samples/snippets/csharp/VS_Snippets_CFX/c_wcfclienttowseservice/cs/wsehttpbinding.cs#3)]
         [!code-vb[c_WCFClientToWSEService#3](../../../../samples/snippets/visualbasic/VS_Snippets_CFX/c_wcfclienttowseservice/vb/wsehttpbinding.vb#3)]  
  
    3.  Переопределение <xref:System.ServiceModel.Channels.Binding.CreateBindingElements%2A> метод для задания свойств привязки.  
  
         В следующем примере кода транспорт, кодирование сообщений и параметры защиты сообщений задаются получением значений свойств `SecurityAssertion` и `MessageProtectionOrder`.  
  
         [!code-csharp[c_WCFClientToWSEService#2](../../../../samples/snippets/csharp/VS_Snippets_CFX/c_wcfclienttowseservice/cs/wsehttpbinding.cs#2)]
         [!code-vb[c_WCFClientToWSEService#2](../../../../samples/snippets/visualbasic/VS_Snippets_CFX/c_wcfclienttowseservice/vb/wsehttpbinding.vb#2)]  
  
3.  В коде клиентского приложения добавьте код для задания свойств привязки.  
  
     В следующем примере кода указывается, что клиент [!INCLUDE[indigo2](../../../../includes/indigo2-md.md)] должен использовать защиту сообщений и проверку подлинности в соответствии с готовым к использованию утверждением безопасности WSE 3.0 `AnonymousForCertificate`. Кроме того, требуются безопасные сеансы и производные ключи.  
  
     [!code-csharp[c_WCFClientToWSEService#4](../../../../samples/snippets/csharp/VS_Snippets_CFX/c_wcfclienttowseservice/cs/client.cs#4)]
     [!code-vb[c_WCFClientToWSEService#4](../../../../samples/snippets/visualbasic/VS_Snippets_CFX/c_wcfclienttowseservice/vb/client.vb#4)]  
  
## <a name="example"></a>Пример  
 В следующем примере кода определяется пользовательская привязка, предоставляющая свойства, соответствующие свойствам готового к использованию утверждения безопасности WSE 3.0. Затем пользовательская привязка с именем `WseHttpBinding` используется для задания свойств привязки для клиента [!INCLUDE[indigo2](../../../../includes/indigo2-md.md)].  
  
  
[!code-csharp[c_WCFClientToWSEService#0](../../../../samples/snippets/csharp/VS_Snippets_CFX/c_wcfclienttowseservice/cs/client.cs#0)]
[!code-vb[c_WCFClientToWSEService#0](../../../../samples/snippets/visualbasic/VS_Snippets_CFX/c_wcfclienttowseservice/vb/client.vb#0)]  
  
## <a name="see-also"></a>См. также  
 <xref:System.ServiceModel.Channels.Binding>   
 [Взаимодействие с WSE](http://msdn.microsoft.com/ru-ru/f6816861-96a0-45f9-8736-8e4e82cd3a41)