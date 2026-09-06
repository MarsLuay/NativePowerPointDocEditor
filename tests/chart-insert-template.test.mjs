import test from "node:test";
import assert from "node:assert/strict";
import { loadChartInsertTemplateModule } from "./helpers/load-plugin-modules.mjs";

test("chart insert template exports correct base64 and xml strings", async () => {
    const {
        CHART_INSERT_FRAME_TEMPLATE,
        CHART_INSERT_CHART_XML,
        CHART_INSERT_CHART_RELS_XML,
        CHART_INSERT_WORKBOOK_BASE64
    } = await loadChartInsertTemplateModule();

    assert.ok(CHART_INSERT_FRAME_TEMPLATE.includes("<p:graphicFrame>"));
    assert.ok(CHART_INSERT_CHART_XML.includes("<c:chartSpace"));
    assert.ok(CHART_INSERT_CHART_RELS_XML.includes("<Relationships"));
    // check if it is valid base64
    const buffer = Buffer.from(CHART_INSERT_WORKBOOK_BASE64, "base64");
    assert.ok(buffer.length > 0);
});
