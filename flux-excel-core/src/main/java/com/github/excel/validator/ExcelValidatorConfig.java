package com.github.excel.validator;

import org.hibernate.validator.HibernateValidator;
import org.hibernate.validator.messageinterpolation.ResourceBundleMessageInterpolator;
import org.hibernate.validator.resourceloading.PlatformResourceBundleLocator;

import jakarta.validation.Validation;
import jakarta.validation.Validator;
import jakarta.validation.ValidatorFactory;

/**
 * Excel 校验
 * @author Vico
 * @create 2022-12-12 18:25
 */
public class ExcelValidatorConfig {
    private static ValidatorFactory validatorFactory = null;

    public static Validator getValidator(){
        if (validatorFactory == null) {
            validatorFactory = Validation.byProvider(HibernateValidator.class)
                    .configure().messageInterpolator(new ResourceBundleMessageInterpolator(
                            new PlatformResourceBundleLocator("messages" )))
                    .failFast(true).buildValidatorFactory();
        }
        return validatorFactory.getValidator();
    }
}
