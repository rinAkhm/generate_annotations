# Данный скрипт разработан для создания документации (аннотация) к РПД.

## Для старта необходимо:
```
git clone https://github.com/rinAkhm/generate_annotations.git
cd generate_annotations

pip install -r requirements.txt
```

## Создать проект в GoogleAPI 
Для этого: 
1. Создать проект в [Управление проектами](https://console.developers.google.com/cloud-resource-manager)
2. Подключить API google drive и google sheets
3. Добавить созданный аккаунт в таблицу google sheets
4. Скачать ключ формата json и добавить в папку проекта

Дополнительную информацию можно почитать [здесь](https://habr.com/ru/post/483302/) 

## Для запуска 
Необходимо установить python3. Для запуска нужно открыть папку с репозитория и прописать
```
python sheeets.py
```

## Примечание. 
Чтобы данные подгружались правильно необходимо заполнять таблицу в данном формате. 
