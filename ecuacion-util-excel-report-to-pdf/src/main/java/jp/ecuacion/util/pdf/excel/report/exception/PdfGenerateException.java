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
package jp.ecuacion.util.pdf.excel.report.exception;

import jp.ecuacion.lib.core.exception.ViolationException;
import jp.ecuacion.lib.core.violation.BusinessViolation;
import jp.ecuacion.lib.core.violation.Violations;
import org.jspecify.annotations.Nullable;

/**
 * Is the common superclass of exceptions thrown when PDF generation fails for a reason that may
 * originate from the given Excel file or from the runtime environment (e.g. a font it depends
 * on), as opposed to a purely technical failure such as an {@code IOException} from the
 * underlying libraries.
 *
 * <p>Each specific failure is represented by one of the concrete subclasses in this package
 * (e.g. {@link SheetNotExistException}, {@link CharacterNotRenderableException}), so callers can
 * {@code catch} the specific case they want to handle differently instead of branching on a
 * {@code messageId} string. Catching this class itself still works for callers that only want
 * to handle "some PDF generation problem" generically.</p>
 */
public abstract class PdfGenerateException extends ViolationException {

  private static final long serialVersionUID = 1L;

  /**
   * Constructs an instance.
   *
   * @param messageId messageId
   * @param messageArgs messageArgs
   */
  protected PdfGenerateException(String messageId, @Nullable Object... messageArgs) {
    super(new Violations().add(new BusinessViolation(messageId, messageArgs)));
  }

  /**
   * Gets messageId.
   *
   * @return messageId
   */
  public String getMessageId() {
    return getViolations().getBusinessViolations().get(0).getMessageId();
  }
}
