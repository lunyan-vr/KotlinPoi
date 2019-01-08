/*
 * Copyright 2019 Naoaki Izumaru
 *
 * Licensed under the MIT
 * You may obtain a copy of the License at
 *
 * https://opensource.org/licenses/mit-license.php
 */
package com.example

import org.apache.logging.log4j.LogManager
import org.apache.poi.hpsf.DocumentSummaryInformation
import org.apache.poi.hpsf.PropertySet
import org.apache.poi.hpsf.SummaryInformation
import org.apache.poi.poifs.filesystem.DirectoryEntry
import org.apache.poi.poifs.filesystem.DocumentEntry
import org.apache.poi.poifs.filesystem.DocumentInputStream
import org.apache.poi.poifs.filesystem.POIFSFileSystem
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory
import org.dom4j.Element
import org.dom4j.io.SAXReader
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException
import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

/** ロガー */
val log = LogManager.getLogger()!!

/** タイトル */
const val EXCEL_PROPERTIES_TITLE = "Title"
/** 件名 */
const val EXCEL_PROPERTIES_SUBJECT = "Subject"
/** タグ */
const val EXCEL_PROPERTIES_KEYWORDS = "Keywords"
/** 分類項目 */
const val EXCEL_PROPERTIES_CATEGORY = "Category"
/** コメント */
const val EXCEL_PROPERTIES_DESCRIPTION = "Description"
/** 作成者 */
const val EXCEL_PROPERTIES_CREATOR = "Creator"
/** 前回保存者 */
const val EXCEL_PROPERTIES_LAST_MODIFIED_BY_USER = "lastModifiedByUser"
/** 会社名 */
const val EXCEL_PROPERTIES_COMPANY = "Company"
/** マネージャ */
const val EXCEL_PROPERTIES_MANAGER = "Manager"

/**
 * 変更するExcelのプロパティを取得し、反映する
 */
fun main(args: Array<String>) {
    log.debug(System.getProperty("user.dir"))
    // 設定ファイル読込
    val document = SAXReader().read(Paths.get(System.getProperty("user.dir"), "excelProperties.xml").toFile())

    // Excelに設定するプロパティの読込
    val nodes = document.selectNodes("Configuration/Properties/Property")
    var newProperties: Map<String, String> = mutableMapOf()
    for (node in nodes) {
        val propertyName = (node as Element).attributeValue("name")
        newProperties += propertyName to node.text
        log.debug("name:$propertyName, text:${node.text}")
    }

    // Excel読込
    val startDirPath = Paths.get(document.selectSingleNode("Configuration/Directory").text)
    val maxDepth = Integer.parseInt("5")
    Files.walk(
        startDirPath,
        maxDepth
    ).forEach { inputPath ->


        try {
            when (File(inputPath.toString()).extension) {
                "xlsx" -> setPropertyXlsx(inputPath, newProperties)
                "xls" -> setPropertyXls(inputPath, newProperties)
            }

        } catch (e: IOException) {
            e.printStackTrace()
        }

    }

}

/**
 * XLS形式のExcelのプロパティを変更する
 */
fun setPropertyXls(inputPath: Path, newProperties: Map<String, String>): Unit = try {
    log.debug("XLS FILENAME : $inputPath")
    val input = FileInputStream(inputPath.toString())

    POIFSFileSystem(input).use { fs ->
        val dirEntry = fs.root as DirectoryEntry
        // 標準のプロパティ
        val docEntrySi = dirEntry.getEntry(
            SummaryInformation.DEFAULT_STREAM_NAME
        ) as DocumentEntry
        val si = SummaryInformation(DocumentInputStream(docEntrySi).use { PropertySet(it) })

        // 追加の標準プロパティ
        val docEntryDsi = dirEntry.getEntry(DocumentSummaryInformation.DEFAULT_STREAM_NAME) as DocumentEntry
        val dsi = DocumentSummaryInformation(DocumentInputStream(docEntryDsi).use { PropertySet(it) })

        newProperties.forEach { k, v ->
            when (k) {
                // タイトル
                EXCEL_PROPERTIES_TITLE -> si.title = v
                // 件名
                EXCEL_PROPERTIES_SUBJECT -> si.subject = v
                // タグ
                EXCEL_PROPERTIES_KEYWORDS -> si.keywords = v
                // 分類項目
                EXCEL_PROPERTIES_CATEGORY -> dsi.category = v
                // コメント
                EXCEL_PROPERTIES_DESCRIPTION -> si.comments = v

                /* 元の場所 */
                // 作成者
                EXCEL_PROPERTIES_CREATOR -> si.author = v
                // 前回保存者
                EXCEL_PROPERTIES_LAST_MODIFIED_BY_USER -> si.lastAuthor = v

                // 会社
                EXCEL_PROPERTIES_COMPANY -> dsi.company = v
                // 部長
                EXCEL_PROPERTIES_MANAGER -> dsi.manager = v
            }
        }
        FileOutputStream(inputPath.toString()).use {
            si.write(dirEntry, SummaryInformation.DEFAULT_STREAM_NAME)
            dsi.write(dirEntry, DocumentSummaryInformation.DEFAULT_STREAM_NAME)
            fs.writeFilesystem(it)
        }

    }


} finally {

}


/**
 * XLSX形式のExcelのプロパティを変更する
 */
fun setPropertyXlsx(inputPath: Path, newProperties: Map<String, String>): Unit = try {
    log.debug("XLSX FILENAME : $inputPath")
    val input = FileInputStream(inputPath.toString())
    XSSFWorkbookFactory.createWorkbook(input).use { workbook ->
        val xmlProps = workbook.properties
        val coreProps = xmlProps.coreProperties
        val extendProps = xmlProps.extendedProperties.underlyingProperties

        newProperties.forEach { k, v ->
            when (k) {
                /*
                 * コアプロパティ
                 */
                // タイトル
                EXCEL_PROPERTIES_TITLE -> coreProps.title = v
                // 件名
                EXCEL_PROPERTIES_SUBJECT -> coreProps.setSubjectProperty(v)
                // タグ
                EXCEL_PROPERTIES_KEYWORDS -> coreProps.keywords = v
                // 分類項目
                EXCEL_PROPERTIES_CATEGORY -> coreProps.category = v
                // コメント
                EXCEL_PROPERTIES_DESCRIPTION -> coreProps.description = v

                /* 元の場所 */
                // 作成者
                EXCEL_PROPERTIES_CREATOR -> coreProps.creator = v
                // 前回保存者
                EXCEL_PROPERTIES_LAST_MODIFIED_BY_USER -> coreProps.lastModifiedByUser = v

                /*
                 * 拡張プロパティ
                 */
                // 会社
                EXCEL_PROPERTIES_COMPANY -> extendProps.company = v
                // 部長
                EXCEL_PROPERTIES_MANAGER -> extendProps.manager = v
            }
        }

        // ファイル保存
        FileOutputStream(inputPath.toString()).use { output ->
            workbook.write(output)
        }

    }

} finally {

}
