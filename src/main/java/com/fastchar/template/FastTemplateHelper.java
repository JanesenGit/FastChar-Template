package com.fastchar.template;


import com.fastchar.core.FastChar;
import com.fastchar.core.FastHandler;
import com.fastchar.core.FastMapWrap;
import com.fastchar.template.interfaces.IFastTemplateData;
import com.fastchar.template.interfaces.IFastTemplateRender;
import com.fastchar.template.provider.FastExcelTemplateRender;
import com.fastchar.template.provider.FastWordTemplateRender;
import com.fastchar.utils.FastHttpURLConnectionUtils;

import java.io.*;
import java.util.List;
import java.util.Map;

/**
 * 模板渲染工具类，模板的变量声明格式：${variable}、${variable.variable}、${variable[0].variable}
 *
 * @author 沈建（Janesen）
 * @date 2021/12/6 16:32
 */
public class FastTemplateHelper {
    static {
        FastChar.getOverrides()
                .add(FastWordTemplateRender.class)
                .add(FastExcelTemplateRender.class);
    }

    /**
     * 获取变量值
     *
     * @param handler 渲染句柄，可注入到变量方法名中
     * @param key     变量标识符
     * @return 值
     */
    public static Object renderData(FastHandler handler, String key) {
        Object inData = handler.get("__data");
        if (inData instanceof Map) {
            Map<?, ?> dataMap = (Map<?, ?>) inData;
            Object value = FastMapWrap.newInstance(dataMap).get("${" + key + "}");
            if (value != null) {
                return value;
            }
        }
        List<IFastTemplateData> iFastTemplateData = FastChar.getOverrides().newInstances(false, IFastTemplateData.class);
        for (IFastTemplateData iFastTemplateDatum : iFastTemplateData) {
            Object data = iFastTemplateDatum.getData(handler, key);
            if (data != null) {
                return data;
            }
        }
        return null;
    }

    /**
     * 渲染模板
     *
     * @param templateFile 模板文件，支持http格式地址
     * @param saveFile     渲染后保存的文件地址
     */
    public static void renderFile(String templateFile, String saveFile) {
        renderFile(new FastHandler(), templateFile, saveFile);
    }

    /**
     * 渲染模板
     *
     * @param data         数据集合
     * @param templateFile 模板文件，支持http格式地址
     * @param saveFile     渲染后保存的文件地址
     */
    public static void renderFile(Map<String, Object> data, String templateFile, String saveFile) {
        FastHandler handler = new FastHandler();
        handler.put("__data", data);
        renderFile(handler, templateFile, saveFile);
    }

    /**
     * 渲染模板
     *
     * @param data         数据集合
     * @param templateFile 模板文件输入流
     * @param saveFile     渲染后保存的文件地址
     */
    public static void renderFile(Map<String, Object> data, InputStream templateFile, String saveFile) {
        FastHandler handler = new FastHandler();
        handler.put("__data", data);
        renderFile(handler, templateFile, saveFile);
    }


    /**
     * 渲染模板
     *
     * @param data         数据集合
     * @param templateFile 模板文件，支持http格式地址
     * @param saveFile     渲染后保存的文件地址
     */
    public static void renderFile(FastHandler handler, Map<String, Object> data, String templateFile, String saveFile) {
        if (handler == null) {
            handler = new FastHandler();
        }
        handler.put("__data", data);
        renderFile(handler, templateFile, saveFile);
    }

    /**
     * 渲染模板
     *
     * @param data         数据集合
     * @param templateFile 模板文件输入流
     * @param saveFile     渲染后保存的文件地址
     */
    public static void renderFile(FastHandler handler, Map<String, Object> data, InputStream templateFile, String saveFile) {
        if (handler == null) {
            handler = new FastHandler();
        }
        handler.put("__data", data);
        renderFile(handler, templateFile, saveFile);
    }

    /**
     * 渲染模板
     *
     * @param handler      渲染句柄，可注入到变量方法名中
     * @param templateFile 模板文件，支持http格式地址
     * @param saveFile     渲染后保存的文件地址
     */
    @SuppressWarnings("IOStreamConstructor")
    public static void renderFile(FastHandler handler, String templateFile, String saveFile) {
        try {
            if (handler == null) {
                handler = new FastHandler();
            }
            InputStream inputStream;
            if (templateFile.startsWith("http:") || templateFile.startsWith("https:")) {
                inputStream = FastHttpURLConnectionUtils.getInputStream(templateFile);
            } else {
                inputStream = new FileInputStream(templateFile);
            }
            FastTemplateHelper.renderFile(handler, inputStream, saveFile);
        } catch (Exception e) {
            FastChar.getLogger().error(FastTemplateHelper.class, e);
        }
    }


    /**
     * 渲染模板
     *
     * @param handler      渲染句柄，可注入到变量方法名中
     * @param templateFile 模板文件输入流
     * @param saveFile     渲染后保存的文件地址
     */
    @SuppressWarnings("IOStreamConstructor")
    public static void renderFile(FastHandler handler, InputStream templateFile, String saveFile) {
        try {
            if (handler == null) {
                handler = new FastHandler();
            }
            File fileObj = new File(saveFile);
            if (!fileObj.getParentFile().exists()) {
                if (!fileObj.getParentFile().mkdirs()) {
                    FastChar.getLogger().error(FastTemplateHelper.class, new RuntimeException(fileObj.getParent() + "创建失败！"));
                }
            }
            OutputStream outputStream = new FileOutputStream(saveFile);
            handler.put("__fileName", fileObj.getName());
            List<IFastTemplateRender> iFastTemplateRenders = FastChar.getOverrides().newInstances(IFastTemplateRender.class);
            for (IFastTemplateRender iFastTemplateRender : iFastTemplateRenders) {
                iFastTemplateRender.onRender(handler, templateFile, outputStream);
            }
        } catch (Exception e) {
            FastChar.getLogger().error(FastTemplateHelper.class, e);
        }
    }

}
