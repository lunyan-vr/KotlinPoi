KotlinPoi
====
## Overview
This program replaces Excel properties (creator etc).

## Description
This program is developed in the following environment.
+ Language : Kotlin
+ IDE : IntelliJ IDEA Community Edition
+ Apache Poi
+ log4j2
+ dom4j

This is developed as a practice Kotlin and IntelliJ.
Items that can be replaced correspond only to the basic ones displayed in Windows Explorer.

## Usage
Corresponding items are as follows.
Properties not listed are not changed.
```xml
<?xml version="1.0" encoding="UTF-8"?>
<Configuration>
    <Directory>D:\WORKSPACE\Excel</Directory>
    <Properties>
        <Property name="Title">タイトル</Property>
        <Property name="Subject">件名</Property>
        <!-- タグ -->
        <Property name="Keywords">タグ</Property>
        <Property name="Category">分類項目</Property>
        <!-- コメント -->
        <Property name="Description">コメント</Property>
        <Property name="Creator">作成者</Property>
        <Property name="lastModifiedByUser">前回保存者</Property>
        <Property name="Company">会社名</Property>
        <Property name="Manager">マネージャ</Property>
    </Properties>
</Configuration>
```

## Licence
[MIT](https://github.com/tcnksm/tool/blob/master/LICENCE)
