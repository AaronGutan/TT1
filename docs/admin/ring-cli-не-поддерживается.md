# Интерфейс командной строки утилиты ring (не поддерживается)

> Source: https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html
> Synced: 2026-07-02 12:28 from `ring CLI (не поддерживается) __ 1C_Enterprise Development Tools Руководство разработчика.html`

---
Важно: Начиная с версии 2024.1 этот интерфейс не
			поддерживается. Используйте интерфейс [1C:EDT
			CLI](https://its.1c.ru/db/edtdoc/content/10608/hdoc).

- [Версии платформы](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D1%81%D0%BF%D0%B8%D1%81%D0%BE%D0%BA-%D0%B2%D0%B5%D1%80%D1%81%D0%B8%D0%B9-%D0%BF%D0%BB%D0%B0%D1%84%D1%82%D0%BE%D1%80%D0%BC%D1%8B)
- [Версии 1C:EDT](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D1%81%D0%BF%D0%B8%D1%81%D0%BE%D0%BA-%D0%B2%D0%B5%D1%80%D1%81%D0%B8%D0%B9-%D0%B5%D0%B4%D1%82)
- [Назначить версию 1C:EDT по умолчанию](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D0%BD%D0%B0%D0%B7%D0%BD%D0%B0%D1%87%D0%B8%D1%82%D1%8C-%D0%B2%D0%B5%D1%80%D1%81%D0%B8%D1%8E-%D0%B5%D0%B4%D1%82-%D0%BF%D0%BE-%D1%83%D0%BC%D0%BE%D0%BB%D1%87%D0%B0%D0%BD%D0%B8%D1%8E)
- [Оптимизировать формат хранения](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D0%BE%D0%BF%D1%82%D0%B8%D0%BC%D0%B8%D0%B7%D0%B8%D1%80%D0%BE%D0%B2%D0%B0%D1%82%D1%8C-%D1%84%D0%BE%D1%80%D0%BC%D0%B0%D1%82-%D1%85%D1%80%D0%B0%D0%BD%D0%B5%D0%BD%D0%B8%D1%8F)
- [Экспорт в XML](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D1%8D%D0%BA%D1%81%D0%BF%D0%BE%D1%80%D1%82-%D0%BF%D1%80%D0%BE%D0%B5%D0%BA%D1%82%D0%B0)
- [Импорт из XML](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D0%B8%D0%BC%D0%BF%D0%BE%D1%80%D1%82-%D0%BA%D0%BE%D0%BD%D1%84%D0%B8%D0%B3%D1%83%D1%80%D0%B0%D1%86%D0%B8%D0%B8)
- [Проверка](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D0%BF%D1%80%D0%BE%D0%B2%D0%B5%D1%80%D0%BA%D0%B0)

При установке 1C:EDT также устанавливается утилита [ring](https://its.1c.ru/db/v8321doc/bookmark/adm/TI000000674), входящая в состав платформы «1С:Предприятие». С ее помощью можно выполнять ряд операций, автоматизирующих ваш процесс разработки. Все команды, предназначенные для работы 1C:EDT находятся в модуле edt этой утилиты.

Стандартно утилита ring находится в каталоге:

- ОС Windows: %ProgramFiles%\1C\1CE\components\1c-enterprise-ring-<версия>-x86_64;
- ОС Linux: /opt/1C/1CE/components/1c-enterprise-ring-<версия>-x86_64;
- ОС macOS: /Applications/1C/1CE/components/1c-enterprise-ring-<версия>-x86_64.

Чтобы получить справку по всем командам интерфейса командной строки, выполните:

ring edt@2021.3.1:x86_64 help
		где 2021.3.1:x86_64 — версия интересующего вас модуля edt.

Получить список всех установленных модулей (с версиями) можно командой:

ring help modules
## Список версий платформы
Чтобы получить список поддерживаемых версий платформы используйте команду platform-versions. Она имеет один параметр:

- --edt-location — каталог, содержащий нужную версию 1C:EDT. Если параметр не задан, то используется либо [версия по умолчанию](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D0%BD%D0%B0%D0%B7%D0%BD%D0%B0%D1%87%D0%B8%D1%82%D1%8C-%D0%B2%D0%B5%D1%80%D1%81%D0%B8%D1%8E-%D0%B5%D0%B4%D1%82-%D0%BF%D0%BE-%D1%83%D0%BC%D0%BE%D0%BB%D1%87%D0%B0%D0%BD%D0%B8%D1%8E), либо [самая ранняя установка](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D1%81%D0%BF%D0%B8%D1%81%D0%BE%D0%BA-%D0%B2%D0%B5%D1%80%D1%81%D0%B8%D0%B9-%D0%B5%D0%B4%D1%82), связанная с модулем edt
Примечание: Для выполнения любого действия с 1C:EDT инициализируется контейнер с OSGi, которому требуется [рабочая область](https://its.1c.ru/db/edtdoc/content/10330/hdoc). Поэтому при выполнении этой команды будет создана временная рабочая область в temp-каталоге пользователя. Она будет удалена после выполнения команды.
Например:

ring edt@2021.3.1:x86_64 platform-versions
## Список версий 1C:EDT
Чтобы получить список установленных версий 1C:EDT используйте команду list из подсистемы locations. Она не имеет параметров.

Например:

ring edt@2021.3.1:x86_64 locations list
## Назначить версию 1C:EDT по умолчанию
Чтобы назначить версию 1C:EDT по умолчанию для данного модуля используйте команду set-default из подсистемы locations.

Она имеет один параметр — путь к установленной версии 1C:EDT. Используйте значение @none чтобы сбросить установку по умолчанию

Например:

ring edt@2021.3.1:x86_64 locations set-default C:\Program Files\1C\1CE\components\1c-edt-2021.3.1+18-x86_64
## Оптимизировать формат хранения данных проекта
Чтобы [оптимизировать формат хранения](https://its.1c.ru/db/edtdoc/content/10159/hdoc/t000159__%D0%BE%D0%BF%D1%82%D0%B8%D0%BC%D0%B8%D0%B7%D0%B0%D1%86%D0%B8%D1%8F-%D1%84%D0%BE%D1%80%D0%BC%D0%B0%D1%82%D0%B0-%D1%85%D1%80%D0%B0%D0%BD%D0%B5%D0%BD%D0%B8%D1%8F) данных проекта используйте команду clean-up-source из подсистемы workspace. Она имеет следующие параметры:

- --workspace-location — обязательный параметр. Каталог [рабочей области](https://its.1c.ru/db/edtdoc/content/10330/hdoc);
- --edt-location — каталог, содержащий нужную версию 1C:EDT. Если параметр не задан, то используется либо [версия по умолчанию](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D0%BD%D0%B0%D0%B7%D0%BD%D0%B0%D1%87%D0%B8%D1%82%D1%8C-%D0%B2%D0%B5%D1%80%D1%81%D0%B8%D1%8E-%D0%B5%D0%B4%D1%82-%D0%BF%D0%BE-%D1%83%D0%BC%D0%BE%D0%BB%D1%87%D0%B0%D0%BD%D0%B8%D1%8E), либо [самая ранняя установка](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D1%81%D0%BF%D0%B8%D1%81%D0%BE%D0%BA-%D0%B2%D0%B5%D1%80%D1%81%D0%B8%D0%B9-%D0%B5%D0%B4%D1%82), связанная с модулем edt;
- --project— каталог проекта, который нужно оптимизировать. Одновременно можно использовать только один параметр: project или project-name;
- --project-name — имя проекта в текущей [рабочей области](https://its.1c.ru/db/edtdoc/content/10330/hdoc), который нужно оптимизировать. Одновременно можно использовать только один аргумент: project или project-name.

Пример:

ring edt@2021.3.1:x86_64 workspace clean-up-source --workspace-location C:\projects\2021.3.0
## Экспорт проекта в xml-выгрузку конфигурации
Чтобы конвертировать проект из файлового представления 1C:EDT в xml-выгрузку конфигурации используйте команду export из подсистемы workspace. Она имеет следующие параметры:

- --workspace-location — обязательный параметр. Каталог [рабочей области](https://its.1c.ru/db/edtdoc/content/10330/hdoc);
- --configuration-files — обязательный параметр. Каталог, в который следует поместить xml-выгрузку конфигурации;
- --edt-location — каталог, содержащий нужную версию 1C:EDT. Если параметр не задан, то используется либо [версия по умолчанию](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D0%BD%D0%B0%D0%B7%D0%BD%D0%B0%D1%87%D0%B8%D1%82%D1%8C-%D0%B2%D0%B5%D1%80%D1%81%D0%B8%D1%8E-%D0%B5%D0%B4%D1%82-%D0%BF%D0%BE-%D1%83%D0%BC%D0%BE%D0%BB%D1%87%D0%B0%D0%BD%D0%B8%D1%8E), либо [самая ранняя установка](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D1%81%D0%BF%D0%B8%D1%81%D0%BE%D0%BA-%D0%B2%D0%B5%D1%80%D1%81%D0%B8%D0%B9-%D0%B5%D0%B4%D1%82), связанная с модулем edt;
- --project — каталог проекта, который нужно экспортировать. Одновременно можно использовать только один аргумент: project или project-name;
- --project-name — имя проекта в текущей [рабочей области](https://its.1c.ru/db/edtdoc/content/10330/hdoc), который следует экспортировать. Одновременно можно использовать только один аргумент: project или project-name.

Пример выполнения:

ring edt@2021.3.1:x86_64 workspace export --project D:/project-1 --configuration-files d:/XML-1/ --workspace-location D:/workspace
## Импорт xml-выгрузки конфигурации в проект
Чтобы конвертировать xml-выгрузку конфигурации в файловое представление 1C:EDT используйте команду import из подсистемы workspace. Она имеет следующие параметры:

- --workspace-location — обязательный параметр. Каталог [рабочей области](https://its.1c.ru/db/edtdoc/content/10330/hdoc);
- --configuration-files — обязательный параметр. Каталог, содержащий xml-выгрузку конфигурации;
- --edt-location — каталог, содержащий нужную версию 1C:EDT. Если параметр не задан, то используется либо [версия по умолчанию](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D0%BD%D0%B0%D0%B7%D0%BD%D0%B0%D1%87%D0%B8%D1%82%D1%8C-%D0%B2%D0%B5%D1%80%D1%81%D0%B8%D1%8E-%D0%B5%D0%B4%D1%82-%D0%BF%D0%BE-%D1%83%D0%BC%D0%BE%D0%BB%D1%87%D0%B0%D0%BD%D0%B8%D1%8E), либо [самая ранняя установка](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D1%81%D0%BF%D0%B8%D1%81%D0%BE%D0%BA-%D0%B2%D0%B5%D1%80%D1%81%D0%B8%D0%B9-%D0%B5%D0%B4%D1%82), связанная с модулем edt;
- --project — каталог проекта, в который следует импортировать проект в формате 1C:EDT. Одновременно можно использовать только один аргумент: project или project-name;
- --project-name — имя проекта в текущей [рабочей области](https://its.1c.ru/db/edtdoc/content/10330/hdoc), в который следует импортировать проект в формате 1C:EDT. Одновременно можно использовать только один аргумент: project или project-name;
- --version — версия платформы «1С:Предприятие 8». Если не указана, то будет подобрана согласно версии xml-выгрузки конфигурации;
- --base-project-name — имя базового проекта. Допустимо только для зависимых проектов.

Примеры выполнения:

ring edt@2021.3.1:x86_64 workspace import --configuration-files d:/XML-1/ --project D:/project-1 --workspace-location D:/workspace
ring edt@2021.3.1:x86_64 workspace import --configuration-files d:/XML-2/ --project D:/project-2 --base-project-name project-1 --workspace-location D:/workspace
ring edt@2021.3.1:x86_64 workspace import --configuration-files d:/XML-2/ --project D:/project-2 --base-project-name project-1 --version 8.3.11 --workspace-location D:/workspace
## Проверка проекта
Чтобы проверить проект используйте команду validate из подсистемы workspace. Она имеет следующие параметры:

- --workspace-location — обязательный параметр. Каталог [рабочей области](https://its.1c.ru/db/edtdoc/content/10330/hdoc);
- --edt-location — каталог, содержащий нужную версию 1C:EDT. Если параметр не задан, то используется либо [версия по умолчанию](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D0%BD%D0%B0%D0%B7%D0%BD%D0%B0%D1%87%D0%B8%D1%82%D1%8C-%D0%B2%D0%B5%D1%80%D1%81%D0%B8%D1%8E-%D0%B5%D0%B4%D1%82-%D0%BF%D0%BE-%D1%83%D0%BC%D0%BE%D0%BB%D1%87%D0%B0%D0%BD%D0%B8%D1%8E), либо [самая ранняя установка](https://its.1c.ru/db/content/edtdoc/src/topics/t000003.html#t000003__%D1%81%D0%BF%D0%B8%D1%81%D0%BE%D0%BA-%D0%B2%D0%B5%D1%80%D1%81%D0%B8%D0%B9-%D0%B5%D0%B4%D1%82), связанная с модулем edt;
- --file — обязательный параметр. Файл для записи результатов валидации в формате TSV. Если файл уже существует, будет выдана ошибка;
- --project-list — список каталогов, откуда загрузить проекты в формате 1C:EDT для проверки. Одновременно можно использовать только один аргумент: project-list или project-name-list;
- --project-name-list — список имен проектов в текущей [рабочей области](https://its.1c.ru/db/edtdoc/content/10330/hdoc), откуда загрузить проекты в формате 1C:EDT для проверки. Одновременно можно использовать только один аргумент: project-list или project-name-list;

Пример выполнения:
ring edt@2021.3.1:x86_64 workspace validate --project-list D:/project-1 D:/project-2 --file D:/validation-result.txt --workspace-location D:/workspace
			

		

	



**На уровень выше:** [Интерфейс командной строки](https://its.1c.ru/db/edtdoc/content/10005/hdoc)