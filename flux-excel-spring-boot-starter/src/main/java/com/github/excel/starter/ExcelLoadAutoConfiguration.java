package com.github.excel.starter;

import com.github.excel.boot.ExcelMetadataRegistry;
import com.github.excel.config.ExcelProperties;
import com.github.excel.exception.ExcelReaderException;
import com.github.excel.read.facade.AbstractReaderTemplateExclude;
import com.github.excel.util.StringUtil;
import com.google.common.base.Throwables;
import com.google.common.collect.Maps;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.write.ExcelLargeListWriter;
import com.github.excel.write.impl.ExcelLargeListWriterImpl;
import lombok.extern.java.Log;
import org.springframework.boot.autoconfigure.AutoConfiguration;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.core.io.Resource;
import org.springframework.core.io.support.ResourcePatternResolver;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel加载自动配置
 */
@AutoConfiguration
@EnableConfigurationProperties({ExcelProperties.class,})
@Log
public class ExcelLoadAutoConfiguration {

	@Autowired
	private ExcelProperties excelProperties ;

	@Autowired(required = false)
	private AbstractReaderTemplateExclude templateExcluder ;

	@Autowired
	private ResourcePatternResolver resourcePatternResolver;

	@Bean
	@ConditionalOnMissingBean
	public ExcelLargeListWriter excelWriter(){
		log.info("==========================");
		log.info("Powered by flux-excel.");
		log.info("==========================");
		ExcelProperties.configMap.put("exportTemplateDir", excelProperties.getExportTemplateDir());
		ExcelProperties.configMap.put("importTemplateDir", excelProperties.getImportTemplateDir());
		ExcelMetadataRegistry.configureTemplateDirectory(excelProperties.getExportTemplateDir(), excelProperties.getImportTemplateDir(), templateExcluder);
		ExcelMetadataRegistry.configureTemplateResolver(this::resolveTemplate);

		return new ExcelLargeListWriterImpl("sheet");
	}

	private List<Map<String, Object>> resolveTemplate(String templateDir, String templateName) {
		List<Map<String, Object>> fileList = new ArrayList<>();
		if (StringUtil.isEmpty(templateDir) || StringUtil.isEmpty(templateName)) {
			return fileList;
		}
		addResource2List("classpath*:" + templateDir + ExcelConstant.FILE_SEPARATOR + templateName, fileList);
		return fileList;
	}

	private void addResource2List(String path, List<Map<String, Object>> files){
		try {
			log.info("load resource from " + path);
			Resource[] resources = resourcePatternResolver.getResources(path);
			if(resources != null){
				for(Resource resource: resources){
					Map<String, Object> map = Maps.newHashMap();
					map.put("name", resource.getFilename());
					log.info("load excel template with name " + resource.getFilename());
					map.put("input", resource.getInputStream());
					map.put("cacheInput", resource.getInputStream());
					files.add(map);
				}
			}
		} catch (IOException e){
			log.info("load resources error with flux-excel, cause:" + Throwables.getStackTraceAsString(e));
			throw new ExcelReaderException("template.load.fail");
		}
	}
}
