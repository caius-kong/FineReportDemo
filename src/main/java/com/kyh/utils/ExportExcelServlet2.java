package com.kyh.utils;

import com.fr.base.FRContext;
import com.fr.base.Parameter;
import com.fr.dav.LocalEnv;
import com.fr.general.ModuleContext;
import com.fr.io.TemplateWorkBookIO;
import com.fr.io.exporter.ExcelExporter;
import com.fr.io.exporter.PageExcelExporter;
import com.fr.main.impl.WorkBook;
import com.fr.main.workbook.PageWorkBook;
import com.fr.report.core.ReportUtils;
import com.fr.report.module.EngineModule;
import com.fr.report.report.PageReport;
import com.fr.stable.PageActor;
import com.fr.stable.WriteActor;
import org.apache.commons.lang3.StringUtils;
import org.joda.time.DateTime;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * Created by kongyunhui on 16/10/24.
 */
public class ExportExcelServlet2 extends javax.servlet.http.HttpServlet {
    protected void doPost(javax.servlet.http.HttpServletRequest request, javax.servlet.http.HttpServletResponse response) throws javax.servlet.ServletException, IOException {
        doGet(request, response);
    }

    protected void doGet(javax.servlet.http.HttpServletRequest request, javax.servlet.http.HttpServletResponse response) throws javax.servlet.ServletException, IOException {
        // 初始化(执行环境)
        fineReportInit();
        try {
            // 模拟request参数
            int number = 100;
            int score = 100;

            // 加载模板
            WorkBook workbook = (WorkBook) TemplateWorkBookIO
                    .readTemplateWorkBook(FRContext.getCurrentEnv(),
                            "2个sheet的excel.cpt");
            // 设置模板参数
            Parameter[] parameters = workbook.getParameters();
            for(Parameter parameter : parameters){
                if(parameter.getName().equals("库存") && number!=0)
                    parameter.setValue(number);
                else if(parameter.getName().equals("积分") && score!=0)
                    parameter.setValue(score);
            }
            // 将结果工作薄导出为.xls文件
            FileOutputStream outputStream = new FileOutputStream(new File("/Users/kongyunhui/Downloads/ExcelExport2"+new DateTime().getMillis()+".xls"));
            ExcelExporter ExcelExport = new ExcelExporter();
            ExcelExport.export(outputStream, workbook.execute(null, new WriteActor()));
            // 结束
            outputStream.close();
            ModuleContext.stopModules();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void fineReportInit(){
        String envpath = "/Users/kongyunhui/Documents/study/workspace/intellij/FineReportDemo/src/main/webapp/WEB-INF";
        FRContext.setCurrentEnv(new LocalEnv(envpath));
        ModuleContext.startModule(EngineModule.class.getName());
    }
}
