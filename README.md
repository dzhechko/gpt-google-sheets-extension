# GPT Помощник для Google Sheets

Расширение для Google Sheets, которое добавляет возможности GPT в ваши таблицы через Google Apps Script.

## 🌟 Основные возможности

- 🤖 Интеграция с OpenAI API
- 📊 Анализ данных в выбранных ячейках
- 📑 Генерация сводок по всей таблице
- ⚙️ Настраиваемые параметры модели
- 🔧 Удобный интерфейс настроек

## 📋 Функциональность

### Меню расширения

В Google Sheets добавляется новое меню "GPT Помощник" со следующими опциями:
- **Показать панель ассистента** - открывает боковую панель для работы с GPT
- **Инструменты**
  - Анализ выбранных ячеек
  - Создать сводку
- **Настройки** - настройка параметров API и модели

### Панель ассистента

Боковая панель предоставляет интерфейс для:
- Ввода произвольных запросов к GPT
- Просмотра ответов
- Быстрой очистки диалога
- Поддержка горячих клавиш (Ctrl + Enter для отправки)

### Анализ данных

Функция "Анализ выбранных ячеек" предоставляет:
- Анализ структуры данных
- Выявление ключевых паттернов и трендов
- Статистические наблюдения
- Обнаружение аномалий и выбросов

### Настройки

Панель настроек позволяет конфигурировать:
- Базовый URL API
- Выбор модели GPT
  - GPT-4
  - GPT-4 Turbo
  - GPT-3.5 Turbo
  - GPT-3.5 Turbo 16K
  - Пользовательская модель
- Температура генерации (0-1)
- Максимальное количество токенов

## 🚀 Установка

1. Откройте Google Sheets
2. Перейдите в меню Расширения → Apps Script
3. Создайте новый проект
4. Скопируйте код из файлов проекта:
   - Code.gs
   - Sidebar.html
   - Settings.html
5. Сохраните проект
6. Обновите таблицу

## ⚙️ Настройка

1. Получите API ключ от OpenAI
2. В таблице выберите "GPT Помощник → Настройки"
3. Введите базовый URL API
4. Выберите модель
5. Настройте параметры генерации
6. Сохраните настройки

## 💡 Использование

### Анализ выбранных ячеек
1. Выделите диапазон ячеек для анализа
2. Выберите "GPT Помощник → Инструменты → Анализ выбранных ячеек"
3. Просмотрите результаты анализа в появившемся окне

### Создание сводки
1. Убедитесь, что таблица содержит данные
2. Выберите "GPT Помощник → Инструменты → Создать сводку"
3. Получите общий обзор данных таблицы

### Произвольные запросы
1. Откройте панель ассистента через "GPT Помощник → Показать панель ассистента"
2. Введите ваш запрос
3. Нажмите "Отправить" или используйте Ctrl + Enter

## 🔒 Безопасность

- API ключ хранится в защищенном хранилище Script Properties
- Все запросы к API выполняются через защищенное соединение
- Поддержка пользовательских URL для использования прокси

## 📝 Примечания

- Для работы требуется действующий API ключ OpenAI
- Количество токенов влияет на стоимость запросов
- При использовании пользовательской модели убедитесь в её совместимости с OpenAI API

## 🤝 Вклад в развитие

Мы приветствуем ваш вклад в развитие проекта! Вы можете:
- Создавать Issue с предложениями улучшений
- Присылать Pull Request с исправлениями
- Делиться опытом использования

## 📄 Лицензия

MIT License
