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

"""Ground truth: justified English text (how much Word compresses a line to fit one more word)."""
import random
import sys
from docx import build, p, r

WORDS = ('the agreement shall be governed by and construed in accordance with laws of Japan each party '
         'hereby irrevocably submits to exclusive jurisdiction a an of in to it is on or if as at no so '
         'confidential information means any technical business financial data disclosed receiving '
         'party may not assign transfer rights obligations without prior written consent termination '
         'notwithstanding foregoing liability limited fees paid during twelve months preceding claim '
         'I we us he she they you service services user users license term terms section sections').split()


def paragraph(rng, count):
    words = [rng.choice(WORDS) for _ in range(count)]
    words[0] = words[0].capitalize()
    return ' '.join(words) + '.'


def body():
    rng = random.Random(20261113)
    out = []
    for i in range(40):
        out.append(p(r('E%02d ' % (i + 1) + paragraph(rng, rng.randint(40, 90)))))
    for i in range(10):
        out.append(p(r('L%02d ' % (i + 1) + paragraph(rng, rng.randint(40, 90))), '<w:jc w:val="left"/>'))
    return ''.join(out)


build(sys.argv[1], body())
