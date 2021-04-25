package com.zhuozhengsoft.samples5;

import org.springframework.util.ResourceUtils;

import java.io.FileNotFoundException;

public class DirUtil {
    public static String getDir ()  {
        String path;
        try {
            path= ResourceUtils.getURL("classpath:").getPath();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            throw new RuntimeException("没有找到项目根目录");
        }
        return path;
    }

}
