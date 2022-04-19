/*
 *      Copyright (c) 2018-2028, DreamLu All rights reserved.
 *
 *  Redistribution and use in source and binary forms, with or without
 *  modification, are permitted provided that the following conditions are met:
 *
 *  Redistributions of source code must retain the above copyright notice,
 *  this list of conditions and the following disclaimer.
 *  Redistributions in binary form must reproduce the above copyright
 *  notice, this list of conditions and the following disclaimer in the
 *  documentation and/or other materials provided with the distribution.
 *  Neither the name of the dreamlu.net developer nor the names of its
 *  contributors may be used to endorse or promote products derived from
 *  this software without specific prior written permission.
 *  Author: DreamLu 卢春梦 (596392912@qq.com)
 */
package com.excel;

import org.springframework.lang.Nullable;

import java.util.stream.Stream;

/**
 * 对象工具类
 *
 * @author L.cm
 */
public class ObjectUtil extends org.springframework.util.ObjectUtils {

	/**
	 * 判断元素不为空
	 * @param obj object
	 * @return boolean
	 */
	public static boolean isNotEmpty(@Nullable Object obj) {
		return !ObjectUtil.isEmpty(obj);
	}

	/**
	 * 有 任意 一个 Blank
	 *
	 * @param obj Object
	 * @return boolean
	 */
	public static boolean isAnyEmpty(final Object... obj) {
		if (ObjectUtil.isEmpty(obj)) {
			return true;
		}
		return Stream.of(obj).anyMatch(ObjectUtil::isEmpty);
	}

	/**
	 * 有 任意 一个非 Blank
	 *
	 * @param obj Object
	 * @return boolean
	 */
	public static boolean isAnyNotEmpty(final Object... obj) {
		if (ObjectUtil.isEmpty(obj)) {
			return false;
		}
		return Stream.of(obj).anyMatch(ObjectUtil::isNotEmpty);
	}

	/**
	 * 判断nul设置默认值
	 * @param defaultValue
	 * @return
	 */
	public static Object defaultIfNull(Object object, Object defaultValue) {
		return null != object ? object : defaultValue;
	}


}
