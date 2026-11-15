#!/bin/bash
#
# Copyright © 2012 ecuacion.jp (info@ecuacion.jp)
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#

# Exports a .docx file to PDF with Microsoft Word for Mac.
# usage: export.sh <docx> <pdf>
# Word is sandboxed, so the files are passed through the Office group container,
# which Word can read and write without asking for file access.
C="$HOME/Library/Group Containers/UBF8T346G9.Office/ecuacion-word-gt"
mkdir -p "$C"
rm -f "$C/in.docx" "$C/out.pdf"
cp "$1" "$C/in.docx"
osascript - "$C/in.docx" "$C/out.pdf" <<'EOF'
on run argv
  with timeout of 90 seconds
    tell application "Microsoft Word"
      open file name (item 1 of argv)
      if (count of documents) is 0 then error "Word could not open the file."
      set d to active document
      save as d file name (item 2 of argv) file format format PDF
      close d saving no
    end tell
  end timeout
end run
EOF
if [ -f "$C/out.pdf" ]; then
  cp "$C/out.pdf" "$2"
  echo "exported: $2"
else
  echo "export failed"
  osascript -e 'tell application "System Events" to tell process "Microsoft Word" to return {value of every static text} of window 1' 2>&1 | cut -c1-300
  exit 1
fi
