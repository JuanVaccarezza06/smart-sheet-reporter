package com.smart_sheet.smart_sheet;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.Getter;
import lombok.Setter;

import java.util.Map;
import java.util.Set;

@Data
@AllArgsConstructor
@Getter
@Setter
public class DataRow {

    private Map<String, String> data;

    public String getValueByName(String columnName) {
        return data.get(columnName);
    }

    public Set<String> getKeys() {
        return data.keySet();
    }
}