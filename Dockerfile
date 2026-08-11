FROM php:8.2-apache

# 1. Installa la libreria libzip e l'estensione PHP zip (fondamentale per leggere i file .xlsx)
RUN apt-get update && apt-get install -y \
    libzip-dev \
    zip \
    && docker-php-ext-install zip

# 2. Copia i file del progetto
COPY . /var/www/html/

# 3. Assegna i permessi corretti all'utente di Apache (www-data)
RUN chown -R www-data:www-data /var/www/html/

EXPOSE 80
