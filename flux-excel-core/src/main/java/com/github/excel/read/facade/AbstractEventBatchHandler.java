package com.github.excel.read.facade;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 批量处理大小
 */
public abstract class AbstractEventBatchHandler<T> {

    /**
     * 批量处理的大小，即result到了这个大小之后才会执行batchExecute方法
     */
    @Getter
    @Setter
    private Integer batchSize;

    @Getter
    @Setter
    private Boolean batchExec;

    public AbstractEventBatchHandler(){
        this(100);
    }

    public AbstractEventBatchHandler(Integer batchSize){
        this.batchSize = batchSize;
        this.batchExec = batchSize != null && batchSize > 0;
    }

    /**
     * 批量处理方法,可批量新增，批量更新
     * @param result
     */
    abstract public void batchExecute(List<T> result);

    public void checks(List<T> result) {

    }

}
