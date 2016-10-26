package com.kyh.utils;

import com.fr.base.FRContext;
import com.fr.base.Parameter;
import com.fr.dav.LocalEnv;
import com.fr.general.ModuleContext;
import com.fr.io.TemplateWorkBookIO;
import com.fr.io.exporter.ExcelExporter;
import com.fr.main.impl.WorkBook;
import com.fr.main.parameter.ReportParameterAttr;
import com.fr.report.module.EngineModule;
import com.fr.stable.WriteActor;
import com.google.common.collect.Lists;
import com.google.common.io.Files;
import org.apache.commons.lang3.StringUtils;
import org.joda.time.DateTime;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by kongyunhui on 16/10/24.
 */
public class ExportExcelServlet extends javax.servlet.http.HttpServlet {
    protected void doPost(javax.servlet.http.HttpServletRequest request, javax.servlet.http.HttpServletResponse response) throws javax.servlet.ServletException, IOException {
        doGet(request, response);
    }

    protected void doGet(javax.servlet.http.HttpServletRequest request, javax.servlet.http.HttpServletResponse response) throws javax.servlet.ServletException, IOException {
        // 初始化(执行环境)
        fineReportInit();
        try {
            // 模拟request参数
            String orderTitle = "";
            String userLastCreateTime = "2016-10-25";
            // 加载模板
            WorkBook workbook = (WorkBook) TemplateWorkBookIO
                    .readTemplateWorkBook(FRContext.getCurrentEnv(),
                            "2016年第一季度用户订单汇总.cpt"); // cpt模板文件不可移动,必须放在reportlets文件夹下
            // 设置模板sql参数 (注意: 参数赋值顺序不是按照sql参数出现的顺序)
            Parameter[] parameters = workbook.getParameters();
            for(Parameter parameter : parameters){
                String params_name = parameter.getName();
                if(params_name.equals("订单类型") && StringUtils.isNotBlank(orderTitle))
                    parameter.setValue(orderTitle);
                else if(params_name.equals("截止日期") && StringUtils.isNotBlank(userLastCreateTime))
                    parameter.setValue(userLastCreateTime);
            }
            // 定义parametermap用于执行报表，将执行后的结果工作薄保存为rworkBook
            // 没仔细看api,不清楚这是干嘛的,但是我不需要~~
//            Map parameterMap = new HashMap();
//            for (Parameter parameter : parameters) {
//                parameterMap.put(parameter.getName(), parameter.getValue());
//            }
            // 将结果工作薄导出为.xls文件
            FileOutputStream outputStream = new FileOutputStream(new File("/Users/kongyunhui/Downloads/ExcelExport"+new DateTime().getMillis()+".xls"));
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
