package com.zkr.web.easyExcel.validation;

import com.alibaba.excel.annotation.ExcelProperty;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.groups.Default;
import java.lang.reflect.Field;
import java.util.Set;
/**
 * @author hanjianyao
 * @email hanjianyao@sinosoft.com.cn
 * @date 2020/11/9 14:26
 */
public class EasyExcelValiHelper {
	// excel正则校验方法EasyExcelValiHelper的写法，该方法会根据实体类中的注解来通过正则表达式判断当前单元格内的数据是否符合标准，例如只能是数字之类的，返回的是检查的错误信息
	private EasyExcelValiHelper(){}

	private static Validator validator = Validation.buildDefaultValidatorFactory().getValidator();

	public static <T> String validateEntity(T obj) throws NoSuchFieldException {
		StringBuilder result = new StringBuilder();
		Set<ConstraintViolation<T>> set = validator.validate(obj, Default.class);
		if (set != null && !set.isEmpty()) {
			for (ConstraintViolation<T> cv : set) {
				Field declaredField = obj.getClass().getDeclaredField(cv.getPropertyPath().toString());
				ExcelProperty annotation = declaredField.getAnnotation(ExcelProperty.class);
				//拼接错误信息，包含当前出错数据的标题名字+错误信息
				result.append(annotation.value()[0]+cv.getMessage()).append(";");
			}
		}
		return result.toString();
	}
}
