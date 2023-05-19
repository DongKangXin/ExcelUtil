package com.Xxx.ExcelUtil.annotation;

import com.Xxx.ExcelUtil.enums.Types;

import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Header {

    String value();

    Types type();
}
