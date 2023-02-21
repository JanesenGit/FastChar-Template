package com.fastchar.template.interfaces;

import com.fastchar.core.FastHandler;

/**
 * 模板数据获取
 * @author 沈建（Janesen）
 * @date 2021/12/7 14:23
 */
public interface IFastTemplateData {

    /**
     * 获取变量值
     * @param handler 渲染句柄
     * @param key 变量key
     * @return 返回值
     */
    Object getData(FastHandler handler, String key);
}
