package com.github.spi;

import java.util.List;

/**
 * @author Vico
 * @create 2023-05-17 11:18
 */
public class FileSearch implements Search{
    @Override
    public List<String> searchDoc(String keyword) {
        System.out.println("文件搜索 "+keyword);
        return null;
    }
}