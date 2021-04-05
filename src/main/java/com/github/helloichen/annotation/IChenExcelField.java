package com.github.helloichen.annotation;

import java.lang.annotation.*;

/**
 * @author Chenwp
 * @since 2021-01-18
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface IChenExcelField {
    /**
     * 标记属性为excel字段
     */
    String value() default "";

    /**
     * 是否为导入字段
     */
    boolean importField() default true;

    /**
     * 是否为导出字段
     */
    boolean exportField() default true;

    /**
     * todo 导出时字段顺寻 待开发。。。
     */
    int order() default 0;
}
