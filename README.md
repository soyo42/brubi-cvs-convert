# Convert xls
## Requirements
 - php8.2 docker
   ```
   php:8.2.6-apache-bullseye
   ```
 - install gd extension
   ```
   apt install libfreetype6-dev libjpeg62-turbo-dev libpng-dev
   docker-php-ext-install gd
   ```
 - install zip extension
   ```
   apt install zlib1g-dev zip libzip-dev
   docker-php-ext-install zip
   ```
 - install composer
   ```
   # https://getcomposer.org/doc/00-intro.md
   php -r "copy('https://getcomposer.org/installer', 'composer-setup.php');"
   php -r "if (hash_file('sha384', 'composer-setup.php') === '55ce33d7678c5a611085589f1f3ddf8b3c52d662cd01d4ba75c0ee0459970c2200a51f492d557530c71c15d8dba01eae') { echo 'Installer verified'; } else { echo 'Installer corrupt'; unlink('composer-setup.php'); } echo PHP_EOL;"
   php composer-setup.php
   php -r "unlink('composer-setup.php');"
   ```
 - install phpoffice/phpspreadsheet via composer
   ```
   create composer.json
   #cat composer.json :
   # {
   #  "require": {
   #      "phpoffice/phpspreadsheet": "^1.28"
   #  },
   #  "config": {
   #      "platform": {
   #          "php": "8.2"
   #      }
   #  }
   # }

   php composer.phar install
   ```
   
