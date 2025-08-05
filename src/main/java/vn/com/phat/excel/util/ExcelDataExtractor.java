/*
 * Copyright 2025 tanphat.1095
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package vn.com.phat.excel.util;


import org.springframework.util.Assert;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Utility class for extracting data from objects for Excel processing.
 * This class provides methods to extract data from a list of objects, a single object, and to extract fields from a class.
 * The extracted data is used for processing Excel files.
 *
 * @author phatlt
 */
public class ExcelDataExtractor {

    private ExcelDataExtractor(){}
    public static<T> List<Map<String, Object>> extractData(List<T> data, Class<T> clazz){
        Assert.notNull(data, "Data must not be null");
        List<Map<String,Object>> mappingData = new ArrayList<>(data.size());// Specify the size of the list to avoid resizing
        for(T dat: data){
            mappingData.add(extractData(dat, clazz));
        }
        return mappingData;
    }

    static <T> Map<String, Object> extractData(T data, Class<T> clazz){
        Map<String, Object> map = new HashMap<>();
        List<Field> fields = extractField(clazz);
        for(Field f : fields){
            try {
                f.setAccessible(true); //NOSONAR
                final String name = f.getName();
                final Object value = f.get(data);
                map.put(name.toUpperCase(), value);
            } catch (IllegalAccessException e) {
                // do nothing
            }
        }
        return map;
    }

    public static <T> Map<String, Field> extractFieldMapping(Class<T> clazz){
        Map<String, Field> mapping = new HashMap<>();
        List<Field> fields = extractField(clazz);
        for(Field f : fields){
            mapping.put(f.getName().toUpperCase(), f);
        }

        return mapping;
    }

    public static List<Field> extractField(Class<?> clazz){
        List<Field> field = new ArrayList<>();
        while(clazz!= null && !clazz.equals(Object.class)){
            Field [] fields = clazz.getDeclaredFields();
            field.addAll(Arrays.asList(fields));
            clazz = clazz.getSuperclass();
        }
        return field;
    }
}
