/*
 * Copyright 2014-2015 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.deepoove.poi.policy;

import com.deepoove.poi.NiceXWPFDocument;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.util.StyleUtils;
import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.Collection;

/**
 * 文本数组
 */
public class TextListRenderPolicy extends AbstractRenderPolicy {


	@Override
	protected boolean validate(Object o) {
		return o !=null && o instanceof Collection;
	}

	@Override
	public void doRender(RunTemplate runTemplate, Object o, XWPFTemplate xwpfTemplate){
		if(!(o instanceof Collection)){
			return;
		}
		NiceXWPFDocument doc = xwpfTemplate.getXWPFDocument();
		XWPFRun run = runTemplate.getRun();

		Collection<Object> collection = (Collection<Object>) o;

		XWPFParagraph paragraph;
		XWPFRun newRun;
		for (Object line : collection) {
			// text
			TextRenderData textRenderData = line instanceof TextRenderData
				? (TextRenderData) line : new TextRenderData(line.toString());
			paragraph = doc.insertNewParagraph(run);
			newRun = paragraph.createRun();
			StyleUtils.styleRun(newRun, run);
			StyleUtils.styleRun(newRun, textRenderData.getStyle());
			newRun.setText(textRenderData.getText());
		}
		// 成功后清除标签
		clearPlaceholder(run);
		IRunBody parent = run.getParent();
		run.removeCarriageReturn();
		if (parent instanceof XWPFParagraph) {
			((XWPFParagraph) parent).removeRun(runTemplate.getRunPos());
			// To do: 更好的列表样式
			// ((XWPFParagraph) parent).setSpacingBetween(0,
			// LineSpacingRule.AUTO);
		}
	}

}
