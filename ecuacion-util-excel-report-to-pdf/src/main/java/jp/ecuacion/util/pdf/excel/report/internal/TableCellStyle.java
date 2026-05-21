/*
 * Copyright © 2012 ecuacion.jp (info@ecuacion.jp)
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
package jp.ecuacion.util.pdf.excel.report.internal;

import java.awt.Color;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.jspecify.annotations.Nullable;

/**
 * Pre-resolved visual style contributed by a table style for a specific cell position.
 *
 * <p>All fields are nullable; {@code null} means "not set by the table style" (leave the
 * cell's own style in effect). Font overrides are only applied for the header row.</p>
 */
record TableCellStyle(
    @Nullable Color fill,
    @Nullable BorderStyle topBorderStyle, @Nullable Color topBorderColor,
    @Nullable BorderStyle bottomBorderStyle, @Nullable Color bottomBorderColor,
    @Nullable BorderStyle leftBorderStyle, @Nullable Color leftBorderColor,
    @Nullable BorderStyle rightBorderStyle, @Nullable Color rightBorderColor,
    @Nullable Color fontColor,
    boolean fontBold
) {}
