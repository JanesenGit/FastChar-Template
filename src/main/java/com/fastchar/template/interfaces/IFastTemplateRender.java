package com.fastchar.template.interfaces;

import com.fastchar.core.FastHandler;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * 模板渲染器
 * @author 沈建（Janesen）
 * @date 2021/12/6 17:08
 */
public interface IFastTemplateRender {

    /**
     * 渲染模板
     *
     * @param handler             渲染句柄，当code==0时将停止继续执行渲染
     * @param templateInputStream 模板文件的输入流
     */
    void onRender(FastHandler handler, InputStream templateInputStream, OutputStream newFileOutStream);

}
