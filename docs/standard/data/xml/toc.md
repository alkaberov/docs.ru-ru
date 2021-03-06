# [XML-документы и данные](index.md)
## [Параметры обработки XML](xml-processing-options.md)
## [Обработка XML-данных в памяти](processing-xml-data-in-memory.md)
### [Обработка XML-данных с помощью модели DOM](process-xml-data-using-the-dom-model.md)
#### [Модель объектов XML-документов (DOM)](xml-document-object-model-dom.md)
#### [Типы XML-узлов](types-of-xml-nodes.md)
#### [Иерархия объектной модели (DOM) XML документа](xml-document-object-model-dom-hierarchy.md)
#### [Сопоставление объектной иерархии с XML-данными](mapping-the-object-hierarchy-to-xml-data.md)
#### [Создание XML-документа](xml-document-creation.md)
#### [Считывание XML-документа в DOM](reading-an-xml-document-into-the-dom.md)
##### [Директивы таблицы стилей, внедренных в документе](style-sheet-directives-embedded-in-a-document.md)
##### [Загрузка данных из средства чтения](load-data-from-a-reader.md)
##### [Пробелы и обработка значащих пробелов при загрузке модели DOM](white-space-and-significant-white-space-handling-when-loading-the-dom.md)
##### [Доступ к атрибутам в модели DOM](accessing-attributes-in-the-dom.md)
##### [Считывание объявлений сущностей и ссылок на сущности в DOM](reading-entity-declarations-and-entity-references-into-the-dom.md)
##### [Ссылки на сущности сохраняются](entity-references-are-preserved.md)
##### [— Ссылки на сущности раскрываются и не сохраняются](entity-references-are-expanded-and-not-preserved.md)
#### [Вставка узлов в XML-документа](inserting-nodes-into-an-xml-document.md)
##### [Создание новых узлов в модели DOM](create-new-nodes-in-the-dom.md)
###### [Создание новых атрибутов для элементов в модели DOM](creating-new-attributes-for-elements-in-the-dom.md)
###### [XML-элемент Проверка имен и атрибутов при создании новых узлов](xml-element-and-attribute-name-verification-when-creating-new-nodes.md)
###### [Создание новых ссылок на сущности](creating-new-entity-references.md)
###### [Влияние пространства имен на раскрытие ссылок на сущности для новых узлов, содержащих элементы и атрибуты](namespace-affect-on-entity-ref-expansion-for-new-nodes.md)
##### [Копирование существующих узлов](copy-existing-nodes.md)
##### [Копирование существующих узлов из одного документа в другой](copying-existing-nodes-from-one-document-to-another.md)
##### [Копирование фрагментов документа](copying-document-fragments.md)
#### [Удаление узлов, содержимого и значений из XML-документа](removing-nodes-content-and-values-from-an-xml-document.md)
##### [Удаление узлов из DOM](removing-nodes-from-the-dom.md)
##### [Удаление атрибутов из узла элемента в DOM](removing-attributes-from-an-element-node-in-the-dom.md)
##### [Удаление содержимого узла в модели DOM](removing-node-content-in-the-dom.md)
#### [Изменение узлов, содержимого и значений в XML-документа](modifying-nodes-content-and-values-in-an-xml-document.md)
#### [Проверка XML-документа в DOM](validating-an-xml-document-in-the-dom.md)
#### [Сохранение и запись документа](saving-and-writing-a-document.md)
#### [Выбор узлов с помощью XPath-навигации](select-nodes-using-xpath-navigation.md)
#### [Разрешение внешних ресурсов](resolving-external-resources.md)
#### [Сравнение объекта с помощью класса XmlNameTable](object-comparison-using-xmlnametable.md)
#### [Коллекции узлов в NamedNodeMap и nodelist](node-collections-in-namednodemaps-and-nodelists.md)
##### [Неупорядоченное извлечение узлов по имени или индексу](unordered-node-retrieval-by-name-or-index.md)
##### [Упорядоченное извлечение узлов по индексу](ordered-node-retrieval-by-index.md)
#### [Динамические обновления объектов nodelist и NamedNodeMap](dynamic-updates-to-nodelists-and-namednodemaps.md)
#### [Поддержка пространств имен в модели DOM](namespace-support-in-the-dom.md)
##### [Пространства имен и DTD в DOM](namespaces-and-dtds-in-the-dom.md)
##### [Изменение деклараций пространств имен в XML-документа](changing-namespace-declarations-in-an-xml-document.md)
##### [Изменение свойств префикса пространства имен](changing-namespace-prefix-properties.md)
#### [Обработка событий в XML-документа с помощью XmlNodeChangedEventArgs](event-handling-in-an-xml-document-using-the-xmlnodechangedeventargs.md)
#### [Расширение модели DOM](extending-the-dom.md)
### [Обработка XML-данных с использованием модели данных XPath](process-xml-data-using-the-xpath-data-model.md)
#### [Чтение XML-данных с помощью XPathDocument и XmlDocument](reading-xml-data-using-xpathdocument-and-xmldocument.md)
#### [Выбор, вычисление и отбор XML-данных с помощью XPathNavigator](selecting-evaluating-and-matching-xml-data-using-xpathnavigator.md)
##### [Выберите XML-данных с помощью XPathNavigator](select-xml-data-using-xpathnavigator.md)
##### [Вычисление выражения XPath с помощью XPathNavigator](evaluate-xpath-expressions-using-xpathnavigator.md)
##### [Соответствие узлов с помощью XPathNavigator](matching-nodes-using-xpathnavigator.md)
##### [Типы узлов, распознаваемые запросами XPath](node-types-recognized-with-xpath-queries.md)
##### [Запросы XPath и пространства имен](xpath-queries-and-namespaces.md)
##### [Скомпилированные выражения XPath](compiled-xpath-expressions.md)
##### [Навигация по пространствам имен XPath](xpath-namespace-navigation.md)
#### [Доступ к данным XML с помощью XPathNavigator](accessing-xml-data-using-xpathnavigator.md)
##### [Навигация в наборе узлов с помощью XPathNavigator](node-set-navigation-using-xpathnavigator.md)
##### [Атрибут и навигации узла пространства имен с помощью XPathNavigator](attribute-and-namespace-node-navigation-using-xpathnavigator.md)
##### [Извлечение XML-данных с помощью XPathNavigator](extract-xml-data-using-xpathnavigator.md)
##### [Доступ к строго типизированным XML-данных с помощью XPathNavigator](accessing-strongly-typed-xml-data-using-xpathnavigator.md)
##### [Определяемые пользователем функции и переменные](user-defined-functions-and-variables.md)
#### [Изменение XML-данных с помощью XPathNavigator](editing-xml-data-using-xpathnavigator.md)
##### [Вставка XML-данных с помощью XPathNavigator](insert-xml-data-using-xpathnavigator.md)
##### [Изменение XML-данных с помощью XPathNavigator](modify-xml-data-using-xpathnavigator.md)
##### [Удаление XML-данных с помощью XPathNavigator](remove-xml-data-using-xpathnavigator.md)
#### [Проверка схемы с помощью XPathNavigator](schema-validation-using-xpathnavigator.md)
### [Обработка XML-данных с помощью LINQ to XML](process-xml-data-using-linq-to-xml.md)
## [Преобразования XSLT](xslt-transformations.md)
### [Использование класса XslCompiledTransform](using-the-xslcompiledtransform-class.md)
#### [Входные данные для класса XslCompiledTransform](inputs-to-the-xslcompiledtransform-class.md)
#### [Параметры вывода в классе XslCompiledTransform](output-options-on-the-xslcompiledtransform-class.md)
#### [Разрешение внешних ресурсов во время обработки XSLT](resolving-external-resources-during-xslt-processing.md)
#### [Расширение таблиц стилей XSLT](extending-xslt-style-sheets.md)
##### [Объекты расширения XSLT](xslt-extension-objects.md)
##### [Параметры XSLT](xslt-parameters.md)
##### [Сценарий с помощью блоков msxsl: script](script-blocks-using-msxsl-script.md)
#### [Устранимые ошибки XSLT](recoverable-xslt-errors.md)
#### [Как: преобразовать фрагмент узла](how-to-transform-a-node-fragment.md)
### [Миграция с класса XslTransform](migrating-from-the-xsltransform-class.md)
#### [Как: перенос кода XslTransform](how-to-migrate-your-xsltransform-code.md)
### [Рекомендации по безопасности XSLT](xslt-security-considerations.md)
### [XSLT-компилятор (xsltc.exe)](xslt-compiler-xsltc-exe.md)
#### [Как: выполнить XSLT-преобразование с помощью сборки](how-to-perform-an-xslt-transformation-by-using-an-assembly.md)
### [XSLT-преобразования с помощью класса XslTransform](xslt-transformations-with-the-xsltransform-class.md)
#### [Реализация поведения по выбору в классе XslTransform](implementation-of-discretionary-behaviors-in-the-xsltransform-class.md)
#### [Реализуемых классом XslTransform XSLT-процессора](xsltransform-class-implements-the-xslt-processor.md)
##### [Выходные данные XslTransform](outputs-from-an-xsltransform.md)
##### [XSLT-преобразования над различными хранилищами](xslt-transformations-over-different-stores.md)
##### [Разрешение внешних таблиц стилей XSLT и документов](resolving-external-xslt-style-sheets-and-documents.md)
##### [XsltArgumentList для параметров таблицы стилей и объекты расширения](xsltargumentlist-for-style-sheet-parameters-and-extension-objects.md)
##### [XSLT таблицы стилей сценариев с помощью < msxsl: script >](xslt-stylesheet-scripting-using-msxsl-script.md)
##### [Поддержка функции msxsl: node-set()](support-for-the-msxsl-node-set-function.md)
##### [Наборы узлов в преобразованиях](node-sets-in-transformations.md)
##### [Фрагмент дерева результатов в преобразованиях](result-tree-fragment-in-transformations.md)
#### [XPathNavigator в преобразованиях](xpathnavigator-in-transformations.md)
#### [XPathNodeIterator в преобразованиях](xpathnodeiterator-in-transformations.md)
#### [Ввод XPathDocument в XslTransform](xpathdocument-input-to-xsltransform.md)
#### [Ввод XmlDataDocument в XslTransform](xmldatadocument-input-to-xsltransform.md)
#### [Ввод XmlDocument в XslTransform](xmldocument-input-to-xsltransform.md)
## [Работа с XML-схем](working-with-xml-schemas.md)
### [Модель объектов схемы XML (SOM)](xml-schema-object-model-som.md)
#### [Общие сведения о модели объектов схемы XML](xml-schema-object-model-overview.md)
#### [Чтение и запись XML-схем](reading-and-writing-xml-schemas.md)
#### [Построение XML-схем](building-xml-schemas.md)
#### [Обход XML-схем](traversing-xml-schemas.md)
#### [Изменение XML-схем](editing-xml-schemas.md)
#### [Включение или импорт XML-схем](including-or-importing-xml-schemas.md)
### [XmlSchemaSet для компиляции схемы](xmlschemaset-for-schema-compilation.md)
#### [Информационный набор после компиляции схемы](post-schema-compilation-infoset.md)
#### [Проверка схемы (XSD) XML с помощью XmlSchemaSet](xml-schema-xsd-validation-with-xmlschemaset.md)
#### [Компиляция схемы XmlSchemaCollection](xmlschemacollection-schema-compilation.md)
##### [Проверка XDR с помощью XmlSchemaCollection](xdr-validation-with-xmlschemacollection.md)
##### [Проверка схемы (XSD) XML с помощью XmlSchemaCollection](xml-schema-xsd-validation-with-xmlschemacollection.md)
### [Принудительная проверка с помощью XmlSchemaValidator](xmlschemavalidator-push-based-validation.md)
### [Определение схемы XML](inferring-an-xml-schema.md)
#### [Выведение схем из XML-документов](inferring-schemas-from-xml-documents.md)
#### [Правила выведения структуры и типов узлов схемы](rules-for-inferring-schema-node-types-and-structure.md)
#### [Правила выведения простых типов](rules-for-inferring-simple-types.md)
## [Интеграция XML с реляционными данными и ADO.NET](xml-integration-with-relational-data-and-adonet.md)
## [Управление пространствами имен в XML-документе](managing-namespaces-in-an-xml-document.md)
## [Поддержка типов в классах System.Xml](type-support-in-the-system-xml-classes.md)
### [Сопоставление типов данных XML с типами среды CLR](mapping-xml-data-types-to-clr-types.md)
### [Примечания по реализации поддержки типов XML](xml-type-support-implementation-notes.md)
### [Преобразование типов данных XML](conversion-of-xml-data-types.md)
#### [Преобразование строк в типы данных .NET Framework](converting-strings-to-dotnet-data-types.md)
#### [Преобразование типов .NET Framework в строки](converting-dotnet-types-to-strings.md)
