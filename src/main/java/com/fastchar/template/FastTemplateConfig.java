package com.fastchar.template;

import com.fastchar.interfaces.IFastConfig;

/**
 * @author 沈建（Janesen）
 * @date 2021/12/6 16:30
 */
public class FastTemplateConfig implements IFastConfig {

    private boolean debug;

    public boolean isDebug() {
        return debug;
    }

    public FastTemplateConfig setDebug(boolean debug) {
        this.debug = debug;
        return this;
    }
}
