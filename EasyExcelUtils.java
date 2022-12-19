package com.zkr.common.easyExcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import org.apache.poi.ss.usermodel.IndexedColors;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.List;

/**
 * @description: easyExcel工具类
 */
public class EasyExcelUtils {

    private EasyExcelUtils(){}


    @SuppressWarnings("rawtypes")
	public static void webWriteExcel(HttpServletResponse response, List objects, Class clazz, String fileName) throws IOException {
        String sheetName = fileName;
        webWriteExcel(response,objects,clazz,fileName,sheetName);
    }


    @SuppressWarnings("rawtypes")
	public static void webWriteExcel(HttpServletResponse response, List objects, Class clazz, String fileName, String sheetName) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        response.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8") + ".xlsx");
        // 头的策略
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        // 背景设置为白
        headWriteCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        // 内容的策略
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        HorizontalCellStyleStrategy horizontalCellStyleStrategy =
                new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
        ServletOutputStream outputStream = response.getOutputStream();
        try {
            EasyExcelFactory.write(outputStream, clazz).autoCloseStream(Boolean.FALSE).registerWriteHandler(horizontalCellStyleStrategy).sheet(sheetName).doWrite(objects);
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    public static void write(OutputStream outputStream,List list,Class clazz,String sheetName){
        try {
            // 头的策略
            WriteCellStyle headWriteCellStyle = new WriteCellStyle();
            // 背景设置为白
            headWriteCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
            // 内容的策略
            WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
            HorizontalCellStyleStrategy horizontalCellStyleStrategy =
                    new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);

            EasyExcel.write(outputStream,clazz).autoCloseStream(Boolean.FALSE).registerWriteHandler(horizontalCellStyleStrategy).sheet(sheetName).doWrite(list);
        } catch (Exception e) {
            e.printStackTrace();
        }


    }




}
