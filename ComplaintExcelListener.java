package com.zkr.web.easyExcel.listener;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.zkr.domain.po.ComplaintInfoDO;
import com.zkr.domain.vo.ExcelCheckErrVo;
import com.zkr.feign.IdgenFeign;
import com.zkr.web.easyExcel.validation.EasyExcelValiHelper;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.util.StringUtils;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


/**
 * @author hanjianyao
 * @email hanjianyao@sinosoft.com.cn
 * @date 2020/11/9 14:26
 */
@Data
@EqualsAndHashCode(callSuper=false)
@Slf4j
public class ComplaintExcelListener<T> extends AnalysisEventListener<T> {
	private IdgenFeign idgenFeign;

	// 成功结果集
	private List<T> successList = new ArrayList<>();

	// 失败结果集
	private List<ExcelCheckErrVo<T>> errList = new ArrayList<>();

	private List<ComplaintInfoDO> complaintInfoDOList = new ArrayList();

	private List<T> tList = new ArrayList<>();

	// excel对象的反射类
	private Class<T> clazz;


	public ComplaintExcelListener(Class<T> clazz,IdgenFeign idgenFeign) {
		this.clazz = clazz;
		this.idgenFeign = idgenFeign;
	}

	@Override
	public void invoke(T t, AnalysisContext analysisContext) {
		log.info("开始读取数据");
		int num = 50;
		String errorMsg = "";
		String regEx = "^\\d{4}(\\-)\\d{1,2}\\1\\d{1,2}$";
		Pattern pattern = Pattern.compile(regEx);
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		ComplaintInfoDO complaintInfoDO = new ComplaintInfoDO();
		try {
			//根据excel数据实体中的javax.validation + 正则表达式来校验excel数据
			errorMsg = EasyExcelValiHelper.validateEntity(t);
			Method setErrorMsg = clazz.getDeclaredMethod("setErrorMsg",String.class);
			if (!StringUtils.isEmpty(errorMsg)){
				ExcelCheckErrVo excelCheckErrVo = new ExcelCheckErrVo(t, errorMsg);
				errList.add(excelCheckErrVo);
			}
			setErrorMsg.invoke(t,errorMsg);
			Method getComplaintDate = clazz.getDeclaredMethod("getComplaintDate");
			Method getComplaintAcceptDate = clazz.getDeclaredMethod("getComplaintAcceptDate");
			Method getResultDate = clazz.getDeclaredMethod("getResultDate");
			String complaintDate = getComplaintDate.invoke(t) + "";
			String complaintAcceptDate = getComplaintAcceptDate.invoke(t) + "";
			String resultDate = getResultDate.invoke(t) + "";

			BeanUtils.copyProperties(t, complaintInfoDO);
			complaintInfoDO.setId(idgenFeign.getId());
			// 匹配是否日期格式是否正确
			Matcher matcher = pattern.matcher(complaintDate);
			if (matcher.matches()){
				complaintInfoDO.setComplaintDate(sdf.parse(complaintDate));
			}
			matcher = pattern.matcher(complaintAcceptDate);
			if (matcher.matches()){
				complaintInfoDO.setComplaintAcceptDate(sdf.parse(complaintAcceptDate));
			}
			matcher = pattern.matcher(resultDate);
			if (matcher.matches()){
				complaintInfoDO.setResultDate(sdf.parse(resultDate));
			}
			complaintInfoDOList.add(complaintInfoDO);
			tList.add(t);
		} catch (Exception e) {
			e.printStackTrace();
			errorMsg = "解析excel数据出错";
		}
		// T tt = JSON.parseObject(JSON.toJSONString(t), clazz);
	}

	@Override
	public void doAfterAllAnalysed(AnalysisContext analysisContext) {


		log.info("读取数据完毕");
	}

	/**
	 * @description: 校验excel头部格式，必须完全匹配
	 * @param headMap 传入excel的头部（第一行数据）数据的index,name
	 * @param context
	 * @throws
	 * @return void
	 * @author zhy
	 * @date 2019/12/24 19:27
	 */
	@Override
	public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
		super.invokeHeadMap(headMap, context);
		if (clazz != null){
			try {
				Map<Integer, String> indexNameMap = getIndexNameMap(clazz);
				Set<Integer> keySet = indexNameMap.keySet();
				for (Integer key : keySet) {
					if (com.alibaba.excel.util.StringUtils.isEmpty(headMap.get(key))){
						throw new ExcelAnalysisException("解析excel出错，请传入正确格式的excel");
					}
					if (!headMap.get(key).equals(indexNameMap.get(key))){
						throw new ExcelAnalysisException("解析excel出错，请传入正确格式的excel");
					}
				}

			} catch (NoSuchFieldException e) {
				e.printStackTrace();
			}
		}
	}

	/**
	 * @description: 获取注解里ExcelProperty的value，用作校验excel
	 * @param clazz
	 * @throws
	 * @return java.util.Map<java.lang.Integer,java.lang.String>
	 * @author zhy
	 * @date 2019/12/24 19:21
	 */
	@SuppressWarnings("rawtypes")
	public Map<Integer,String> getIndexNameMap(Class clazz) throws NoSuchFieldException {
		Map<Integer,String> result = new HashMap<>();
		Field field;
		Field[] fields=clazz.getDeclaredFields();
		for (int i = 0; i <fields.length ; i++) {
			field=clazz.getDeclaredField(fields[i].getName());
			field.setAccessible(true);
			ExcelProperty excelProperty=field.getAnnotation(ExcelProperty.class);
			if(excelProperty!=null){
				int index = excelProperty.index();
				String[] values = excelProperty.value();
				StringBuilder value = new StringBuilder();
				for (String v : values) {
					value.append(v);
				}
				result.put(index,value.toString());
			}
		}
		return result;
	}
}
