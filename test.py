# -*- coding: utf-8 -*-
ss = (
    (u'史丹单', 10),
    (u'史丹单丹丹t', 20),
)

for s in ss:

    str = u'{0:<20}, {1:<10}'.format(s[0], s[1], ord(12288))
    print str