FROM php:8.2.6-apache-bullseye


RUN usermod -u 1000 www-data
RUN groupmod -g 1000 www-data

RUN apt update
RUN apt install -y less

# install gd extension
RUN apt install -y libfreetype6-dev libjpeg62-turbo-dev libpng-dev
RUN docker-php-ext-install gd

# install zip extension
RUN apt install -y zlib1g-dev zip libzip-dev
RUN docker-php-ext-install zip

# install composer
# https://getcomposer.org/doc/00-intro.md
COPY install-composer.sh /tmp

USER www-data
WORKDIR /var/www/html

