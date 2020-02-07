package com.gw.test;

import com.gw.service.ExcelUtil;
import com.gw.util.MKSCommand;
import com.mks.api.response.APIException;

public class TestParse {

	public static void main(String[] args) {
		MKSCommand m = new MKSCommand();
		try {
			m.viewIssue("21203",true);
		} catch (APIException e) {
			e.printStackTrace();
		}

	}
}