CREATE TABLE access(
    access_id int auto_increment primary key,
    classroom int NOT NULL,
    access_time DATETIME COLLATE utf8_unicode_ci NOT NULL
);