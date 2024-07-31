# EcoAuto
A RubberduckVBA ultimate experiment.

Цель проекта — перенос из Excel в Word. Наполнение шаблонов документов по результатам межлабораторных сличительных испытаний.

## Постановка задачи

Требуется заполнить данными от 1 до i документов Word.

Данные разделены по типу:
- статичные данные, хранимые в ранее подготовленных файлах;
- данные, получаемые при вводе пользователя;
- данные рабочего листа и вспомогательных книг Excel.

### Статичные данные

Количество файлов с данными от 1 до j.

Данные хранятся в файлах с расширениями:
- *.doc(x);
- *.txt.

Пример каталога, хранящего файлы со статичными данными:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/1%20--%20Static%20Data%20Files.png)

Наименование файла является ключом для поиска в заполняемом шаблоне.
Тело файла полностью переносится в заполняемый шаблон.
При наличии k повторений ключа в шаблоне прозводится k замен.

### Данные, вводимые пользователем

Сведения имеют фиксированное количество:
- Сведения о применяемом субподряде;
- Номер объекта для контроля;
- Сведения об исполнителе;
- Сведения о выбранных участниках испытаний.

#### Сведения о субподряде

Данные о субподряде хранятся в индивидуальных *.txt-файлах.
Наполнение нескольких файлов данных о субподряде комбинируется в единый блок данных и переносится в шаблон как единая строка.
В случае отсутствия выбора хотя бы одного файла данных о субподряде, производится заполненые данными из файла по умолчанию.

Пример каталога, хранящего файлы с данными субподряда:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/2%20--%20Subcontract%20Files.png)

#### Номер объекта для контроля и исполнитель

Номер объекта для контроля необходим для получения из вспомогательной таблицы данных о диапазоне измеряемой единицы величины.
Указанный параметр не передаётся в заполняемый шаблон напрямую, но необходим для заполнения таблиц сравнительных результатов сличений.

Таблица может содержать от 1 до k номеров объектов для контроля и от 1 до m строк единиц величин.

Пример наполнения таблицы с данными объектов для контроля:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/3%20--%20ControlObject%20Table.png)

Сведения об исполнителе передаются в шаблон по фиксированному ключу.

#### Сведения об участниках испытаний

Выбор участников определяет дальнейшее наполнение таблиц результатов испытаний.
Сведения об участниках испытаний являются результатом наложения фильтра на данные рабочего листа Excel.

### Данные рабочего листа Excel

#### Строки участников испытаний

Лист Excel содержит от 1 до n строк с данными участников испытаний.

Каждая строка листа содержит базовые данные участника:
- идентификационный номер;
- номер тура (группы участников) испытаний;
- организационный тип участника;
- наименование организации участника.

Пример блока базовых данных участников испытаний:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/4%20--%20Participant%20Base%20Data.png)

#### Блоки сличения в единице измерений

Каждая строка листа содержит от 1 до m блоков с фиксированным количеством столбцов внутри блока.

Пример блока сравнительных данных сличений в единице измерений:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/5%20--%20Comparison%20Block%20Data.png)

Каждый блок содержит данные испытаний, общие для всех участников блока:
- наименование единицы величины, в которой производились испытания;
- субнаименование единицы величины;
- единица измерений величины;
- эталонное (приписанное) значение единицы величины;
- неопределённость измерения эталонного (приписанного) значения единицы величины.

В случае расхождения данных эталонного (приписанного) значения среди участников блока:
- появляется уведомление пользователя:
- выделяются цветом ячейки, содержащие некорректные значения;
- меню переноса не загружается до устранения несоответствия значений.

Пример уведомления пользователя:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/6%20--%20Incorrect%20Reference%20Notify.png)

Пример выделения цветом ячеек с некорректными данными:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/7%20--%20Incorrect%20Reference%20Mark.png)

Каждый блок содержит данные испытаний, индивидуальные для конкретного участника:
- результат измерения единицы величины;
- неопределённость измерения единицы величины;
- En-критерий оценки соответствия результата участника требованиям испытаний.

Возможна ситуация, когда для конкретного участника отсутствуют результаты измерений в конкретном блоке единицы величины.

#### Округление числовых значений

При чтении в память с листа числовые данные подвергаются округлению.

Значение неопределённости округляется согласно Приложению Е ГОСТ Р 8.736–2011.
Если первая значащая цифра 1, 2 или 3, то округление производится до разряда второй значащей цифры.
Иначе округление производится до разряда первой значащей цифры.

Значение измеренной величины округляется до разряда значения неопределённости.

Значение En-критерия всегда округляется до сотых.

#### Таблицы конкретного участника испытаний

Для каждого участника испытаний должно быть сформировано три типа таблиц:
- таблица измеряемых единиц величин;
- таблица эталонных (приписанных) значений;
- таблица интерпретации результатов испытаний.

Каждая строка таблицы содержит значение участника для каждого блока единиц величин.
Если в блоке отсутствует результат измерений, строка не добавляется в итоговую таблицу.

Шаблон таблицы измеренных величин:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/8%20--%20Measured%20Malues%20Table%20Template.png)

Пример заполнения таблицы измеренных величин:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/9%20--%20Measured%20Values%20Table%20Fill%20Example.png)

Шаблон таблицы результатов участника испытаний:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/10%20--%20Result%20Evaluation%20Table%20Template.png)

Пример заполнения таблицы результатов участника испытаний:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/11%20--%20Result%20Evaluatuion%20Table%20Fill%20Example.png)

Каждая таблица формируется в отдельном файле временного каталога один раз для повторного использования в шаблонах i раз.

#### Таблицы сравнительных результатов сличений выбранных участников испытаний

Для каждого участника формируется от 1 до m таблиц сравнительных результатов испытаний.

Каждая таблица содержит от 1 до n строк с результатами участников.
Строка с данными текущего участника выделяется жирным шрифтом.
Строки с данными других участников не раскрывают их наименование.

Если для текущего участника отсутствуют результаты измерений по конкретному блоку, таблица не будет сформирована.
Если текущий участник единственный в туре блока, таблица не будет сформирована.
Если сторонний участник в туре блока не имеет результатов измерений, строка с его данными не попадает в таблицу.

Шаблон таблицы сравнительных результатов испытаний:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/12%20--%20Comparison%20Table%20Template.png)

Пример заполнения таблиц сравнительных результатов испытаний:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/13%20--%20Comparison%20Table%20Filled%20File%20Content%20Example.png)

Группа таблиц формируется в отдельном файле временного каталога для каждого участника один раз для повторного использования в шаблонах i раз.

#### Диаграммы сравнительных результатов сличений выбранных участников испытаний

Для каждого участника формируется от 1 до m диаграмм сравнительных результатов испытаний.

Каждая диаграмма содержит от 1 до n точек с результатами участников.
Точка с данными текущего участника выделяется специальным маркером.

Если для текущего участника отсутствуют результаты измерений по конкретному блоку, диаграмма не будет сформирована.
Если текущий участник единственный в туре блока, диаграмма не будет сформирована.
Если сторонний участник в туре блока не имеет результатов измерений, точка с его данными не попадает на диаграмму.

Шаблон листа Excel, содержащий диаграмму сравнительных результатов испытаний:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/14%20--%20Comparison%20Chart%20Excel%20Template.png)

Шаблон диаграммы блока единицы величины в документе Word:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/16%20--%20Comparison%20Chart%20Range%20Template.png)

Пример заполнения шаблона листа Excel данными сравнительных испытаний:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/15%20--%20Comparison%20Chart%20Excel%20Fill%20Example.png)

Пример наполнения документа Word диаграммами сравнительных результатов испытаний:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/17%20--%20Comparison%20Charts%20Filled%20File%20Content%20Example.png)

Группа диаграмм формируется в отдельном файле временного каталога для каждого участника один раз для повторного использования в шаблонах i раз.

# UI 

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/18%20--%20UI%20is%20like%20a%20joke.png)

## Ribbon

При запуске книги-стартера Excel происходит запуск надстройки проекта.
На вкладке «Главная» справа появляется группа «EcoAuto» с дополнительными элементами управления.

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/19%20--%20Ribbon%20Buttons.png)

## Main UserForm

Вид на пользовательскую форму ввода данных при первом запуске:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/20%20--%20Main%20User%20Form%20Default%20View.png)

### Статус-бар и статусы валидации

Назначение: 
- отображение статуса валидации компонентов приложения;
- указание пользователю на необходимость заполнить обязательные поля.

Предусмотрено три статуса валидации:
- Полная готовность к переносу: все компоненты обнаружены;
- Частичная возможность переноса: часть компонентов недоступна, перенос возможен в ограниченном объёме;
- Невозможность переноса: недоступны шаблоны, в которые требуется перенести данные.

Вид на пользовательскую форму при частичной возможности переноса (слева) и невозможности переноса (справа): 

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/21%20--%20Main%20User%20Form%20Partial%20Transfer%20Available.png)
![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/22%20--%20Main%20User%20Form%20Transfer%20Unavailable.png)

### Обязательные к заполнению поля

В случае попытки начать перенос при отсутствующих в обязательных полях данных статус-бар отображает соответствующее уведомление, а также подсвечивается целевое поле:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/23%20--%20Input%20Required%20--%20ControlObjectNumber.png)
![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/24%20--%20Input%20Required%20--%20Respondent.png)
![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/25%20--%20Input%20Required%20--%20Participants.png)

### Кеширование вводимых значений

Данные представленных полей будут сохранены в файл кеша по окончании работы с формой:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/26%20--%20Cache%20Fields.png)

Приложение запомнит, какой субподряд и объект для котроля выбрал ранее Полиграф Полиграфович.

### Обработка KeyDown

Элементы формы поддерживают хоткеи:
- нажатие Esc тождественно нажатию на кнопку «Отмена»;
- нажатие Enter тождественно нажатию на кнопку «Перенос».

## Help UserForm

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/28%20--%20Help%20UserForm.png)

Запуск формы возможен двумя способами:
- с Ribbon-ленты;
- с формы ввода пользовательских данных.

При запуске с формы ввода пользовательских данных, после завершения работы с формой справки прозводится повторный вызов формы ввода пользовательских данных.

### Обработка KeyDown

Элементы формы поддерживают хоткеи: нажатие Esc и Enter тождественно нажатию на кнопку «Закрыть».

## Прогресс-бар

Возможности прогресс-бара:
- отображение наименования этапа выполнения;
- отображение шкалы общего состояния выполнения;
- отображения % общего состояния выполнения;
- отображение пояснения для конкретного этапа выполнения.

Особенность работы с прогресс-баром — необходимость корректно оценить суммарное количество этапов вычисления в самом начале работы приложения, чтобы в любом случае выполнение завершалось на 100 %.

Примеры отображения прогресса выполнения:

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/29%20--%20ProgressBar%20--%20Connectiong%20To%20Word.png)

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/30%20--%20ProgressBar%20--%20ComparisonCharts%20Fill.png)

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/31%20--%20ProgressBar%20--%20Results%20Table%20Filling.png)

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/32%20--%20ProgressBar%20--%20Fill%20by%20Key%20Template%202.png)

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/33%20--%20ProgressBar%20--%20SaveResults%20Example.png)

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/34%20--%20ProgressBar%20--%20Final%20Stage.png)

# Особенности архитектуры проекта

![Title](https://github.com/akolodka/EcoAuto/blob/main/resources/35%20--%20First%20Step%20of%20each%20project.jpg)

## Этапы выполнения

Общий порядок выполнения:
- инициализация компонентов, получение исходной модели для переноса;
- дополнение модели переноса данными, введёнными пользователем через форму;
- подготовка вспомогательных файлов с данными для переноса;
- перенос пар ключ-значение статичных данных;
- перенос пар ключ-значение введённых пользователем данных;
- сохранение промежуточных шаблонов результатов испытаний;
- перенос пар ключ-значение данных вспомогательных файлов участников испытаний (таблицы, диаграммы);
- перенос пар ключ-значение базовых данных участников испытаний (наименование, тип);
- сохранение результатов индивидуальных участников.

## Пары ключ-значение

Для переноса данных создаётся объект-контейнер типа IKeyValuePair, хранящий два объекта:
- ITransferKey;
- ITransferValue.

### Ключ ITransferKey

Объект ITransferKey является обёрткой String с той разницей, что при создании проверяется наличие атрибута — фигурных скобок.
При создании ключа из «небезопасного» значения всегда формируется объект, содержащий «безопасное» значение.

При передаче аргумента «key» будет сформирован ключ «{key}».

### Значение ITransferValue

Объект ITransferValue может содержать типы значений:
- простой текст в формате String;
- диапазон документа Word типа IWordRange;
- диаграмма IExcelChart.

При создании объекта ITransferValue также производится контроль и формирование «безопасного» значения.
Для простого текста выполняется удаление пустых символов (ASCII < 32) в начале и конце строки.
Для объектов IWordRange и IExcelChart не выполняется преобразований (возможны потери данных).

### Передача значения в документ

Для каждого типа значения ITransferValue реализована индивидуальная стратегия передачи в документ:
- для простого текста выполняется вставка в диапазон документа назначения по «якорю» из буфера памяти;
- для значения типа IWordRange выполняется копирование диапазона исходного документа и вставка этого диапазона в диапазон документа назначения по «якорю» с сохраненим форматирования;
- для значения типа IExcelChart выполняется копирование объекта-диаграммы рабочего листа и вставка в диапазон документа назначения по «якорю».

«Якорем» является ключ в тексте документа.

## Обращение к Word из Excel

Реализованы два способа начала работы с Word:
- подключение к существующему процессу Word;
- создание нового процесса Word, если не найден существующий процесс;

### Подключение к процессу Word

Подключение к существующему процессу выполняется быстрее, чем создание нового процесса.

Особенности: необходимо сформировать список ранее открытых документов Word, чтобы не закрыть эти документы после выполнения переноса.

### Создание нового процесса

Самый простой в реализации способ: новый процесс для переноса, после выполнения которого принутельное завершение с закрытием всех документов, через него открытых.

Особенности: требуется механизм выгрузки созданного процесса из памяти в случае возникновения ошибки выполнения.

## Обработка ошибок

При возникновении ошибки выполнения в процессе переноса предусмотрен механизм принудительной выгрузки временных документов, а также искусственно созданного процесса Word.

## Отображение в Проводнике каталога по результатам переноса

Целевой каталог может быть открыт ранее, и применение Shell или FileSystemObject приведёт к множественному открытию одноимённых окон Проводника, что создаёт неудобства конечному пользователю.

Для решения указанной проблемы применены функции Win32Api:
- IsIconic Lib "user32.dll";
- ShowWindow Lib "user32".

В случае обнаружения окна в Проводнике происходит перенос фокуса на окно без повторного открытия.

# Заключение

Статистика StackOverflow неумолимо указывает, что VBA — один из худших языков для разработки программных продуктов на 2024 год.

И этому способствуют отсутствие наследования, конструктора в классах, невозможность инициализации переменной при её объявлении и нехватка волшебного LINQ.

Однако при работе с VBA возможно применять ООП-подход, реализовать принципы SOLID, DRY, KISS и конструировать [чистый код](https://www.litres.ru/book/robert-s-martin/chistyy-kod-sozdanie-analiz-i-refaktoring-pdf-epub-6444478/). 

Как оказалось, композиция + интерфейсы могут компенсировать наследование.
Фабрики + [PredeclairedID](https://rubberduckvba.blog/2018/04/24/factories-parameterized-object-initialization/) могут компенсировать конструкторы.
Возможности для инкапсуляции, полиморфизма и абстракции более чем достаточны.

[Mathieu Guindon](https://github.com/retailcoder) верно [отмечал](https://rubberduckvba.blog/2019/04/10/whats-wrong-with-vba/), что проблема VBA не в языке, проблема в VBE IDE. 

[Юнит-тесты](https://rubberduckvba.blog/2017/10/19/how-to-unit-test-vba-code/) (хотя бы базовые, не говоря про TDD) и возможность каталогизировать структуру проекта средствами [Rubberduck](https://rubberduckvba.com/) кардинально меняют процесс разработки.
Mathieu Guindon действительно удалось создать свой reSharper, де-факто необходимый инструмент для любого VBA-разработчика. 

VBA — плохой язык? Определённо, нет. Его просто необходимо [правильно](https://ko-fi.com/s/d91bfd610c) готовить.
