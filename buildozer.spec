[app]

# (str) Назва вашого застосунку
title = Газовий Калькулятор ГРМУ

# (str) Назва пакета (без пробілів і спецсимволів)
package.name = gascalculator

# (str) Домен вашої організації
package.domain = org.grmu

# (str) Вихідний код лежить у тій самій папці
source.dir = .

# (list) Розширення файлів, які треба включити в збірку
source.include_exts = py,png,kv,atlas

# (list) Всі необхідні Python-залежності (актуальні сумісні пакети)
requirements = python3,kivy,kivymd==1.1.1,python-docx,setuptools

# (str) Використання актуальної гілки python-for-android для сумісності з новими середовищами
p4a.branch = master

# (str) Версія застосунку
version = 1.0

# (list) Орієнтація екрана (вертикальна)
orientation = portrait

# (bool) Повноекранний режим
fullscreen = 0

# (list) Дозволи Android
android.permissions = WRITE_EXTERNAL_STORAGE, READ_EXTERNAL_STORAGE, MANAGE_EXTERNAL_STORAGE

# (int) Мінімальна та цільова версії Android API
android.api = 33
android.minapi = 21
android.ndk = 25b

# (bool) Дозвіл на використання застарілих методів збереження файлів
android.requestLegacyExternalStorage = true

# (list) Архітектури процесорів
android.archs = arm64-v8a, armeabi-v7a

[buildozer]

# (int) Рівень логування (2 - детальний)
log_level = 2

# (str) Шлях до папки з результатами збірки
bin_dir = ./bin
android.accept_sdk_license = True
