package com.fastchar.template.info;

import java.util.List;

public class FastWordTableInfo {

    private List<String> titles;

    private List<List<Object>> values;

    public List<String> getTitles() {
        return titles;
    }

    public FastWordTableInfo setTitles(List<String> titles) {
        this.titles = titles;
        return this;
    }

    public List<List<Object>> getValues() {
        return values;
    }

    public FastWordTableInfo setValues(List<List<Object>> values) {
        this.values = values;
        return this;
    }
}
