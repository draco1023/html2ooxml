/*
 * Copyright 2016 - 2021 Draco, https://github.com/draco1023
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.ddr.poi.html;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import org.ddr.poi.FileReader;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

class HtmlRenderTableTest {

    @Test
    void doRender() throws IOException {
        HtmlRenderPolicy htmlRenderPolicy = new HtmlRenderPolicy();
        Configure configure = Configure.builder()
                .bind("text", htmlRenderPolicy)
                .build();
        Map<String, Object> data = new HashMap<>();
        data.put("text", FileReader.readFile("/4.html"));

        try (InputStream inputStream = HtmlRenderPolicyTest.class.getResourceAsStream("/4.docx")) {
            XWPFTemplate.compile(inputStream, configure).render(data).writeToFile("4_out.docx");
        }
    }

}