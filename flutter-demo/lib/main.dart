import 'package:flutter/material.dart';

void main() => runApp(const McrApp());

// ---------- रंग ----------
const kBlue = Color(0xFF1D4ED8);
const kBg = Color(0xFFF4F6FA);
const kText = Color(0xFF111827);
const kMuted = Color(0xFF6B7280);
const kLine = Color(0xFFE5E7EB);

class Bucket {
  final String key, title, sub;
  final Color color, tint;
  const Bucket(this.key, this.title, this.sub, this.color, this.tint);
}

const buckets = [
  Bucket('y1', '1 वर्षावरील', '365+ दिवस', Color(0xFFDC2626), Color(0xFFFEF2F2)),
  Bucket('m6', '6 ते 12 महिने', '181–365 दिवस', Color(0xFFEA580C), Color(0xFFFFF7ED)),
  Bucket('m3', '3 ते 6 महिने', '91–180 दिवस', Color(0xFFCA8A04), Color(0xFFFEFCE8)),
  Bucket('m0', '3 महिन्यांखालील', '0–90 दिवस', Color(0xFF16A34A), Color(0xFFF0FDF4)),
];

Bucket bucketFor(int days) {
  if (days > 365) return buckets[0];
  if (days > 180) return buckets[1];
  if (days > 90) return buckets[2];
  return buckets[3];
}

// ---------- नमुना माहिती ----------
class Case {
  final String crNo, ps, head, kalam, io, complainant, stage;
  final DateTime regDate;
  final int accused, arrested;
  Case(this.crNo, this.ps, this.head, this.kalam, this.io, this.complainant,
      this.stage, this.regDate, this.accused, this.arrested);
  int get days => DateTime.now().difference(regDate).inDays;
  Bucket get bucket => bucketFor(days);
}

final List<Case> sampleCases = _makeCases();

List<Case> _makeCases() {
  const ps = ['सावनेर', 'काटोल', 'नरखेड', 'कळमेश्वर', 'खापा', 'केळवद'];
  const heads = [
    ['घरफोडी', 'BNS 331(4), 305'],
    ['मोटारसायकल चोरी', 'BNS 303(2)'],
    ['जबरी चोरी', 'BNS 309(4)'],
    ['फसवणूक', 'BNS 318(4)'],
    ['दुखापत', 'BNS 115(2), 352'],
    ['विनयभंग', 'BNS 74'],
    ['सायबर फसवणूक', 'BNS 318(4), IT 66D'],
    ['अपघात', 'BNS 281, 125(a)'],
  ];
  const io = ['पोउपनि आर. पाटील', 'सपोनि एस. देशमुख', 'पोउपनि के. वानखेडे', 'पोहवा एम. ठाकरे', 'पोउपनि ए. शेख'];
  const comp = ['रमेश बोरकर', 'सुनीता गायकवाड', 'प्रकाश राऊत', 'अनिल मेश्राम', 'कविता ढोके', 'विजय कावळे', 'शबाना पठाण'];
  const stages = ['तपासावर', 'आरोपी शोध सुरू', 'जप्ती बाकी', 'CA अहवाल प्रलंबित', 'दोषारोपपत्र तयारी', 'साक्षीदार जबाब बाकी'];
  const ages = [
    540, 470, 410, 388, 372, 330, 300, 268, 240, 222, 205, 190, 176, 160, 150,
    138, 121, 110, 98, 95, 84, 70, 61, 52, 45, 38, 30, 22, 15, 9, 6, 3
  ];
  final now = DateTime.now();
  final out = <Case>[];
  for (var i = 0; i < ages.length; i++) {
    final d = now.subtract(Duration(days: ages[i]));
    final h = heads[(i * 3) % heads.length];
    final acc = 1 + (i % 3);
    out.add(Case(
      '${100 + (i * 37) % 800}/${d.year}',
      ps[i % ps.length],
      h[0],
      h[1],
      io[(i * 2) % io.length],
      comp[i % comp.length],
      stages[(i * 5) % stages.length],
      d,
      acc,
      i % 4 == 0 ? acc : (i % 3 == 0 ? 1 : 0),
    ));
  }
  return out;
}

String fmtDate(DateTime d) =>
    '${d.day.toString().padLeft(2, '0')}-${d.month.toString().padLeft(2, '0')}-${d.year}';

// ---------- अॅप ----------
class McrApp extends StatelessWidget {
  const McrApp({super.key});
  @override
  Widget build(BuildContext context) {
    final base = ThemeData(
      useMaterial3: true,
      fontFamily: 'NotoDeva',
      colorScheme: ColorScheme.fromSeed(seedColor: kBlue, primary: kBlue, surface: Colors.white),
      scaffoldBackgroundColor: kBg,
    );
    return MaterialApp(
      title: 'MCR प्रलंबित गुन्हे',
      debugShowCheckedModeBanner: false,
      theme: base.copyWith(
        appBarTheme: const AppBarTheme(
          backgroundColor: Colors.white,
          foregroundColor: kText,
          elevation: 0,
          scrolledUnderElevation: 0.5,
          titleTextStyle: TextStyle(fontFamily: 'NotoDeva', fontSize: 18, fontWeight: FontWeight.w700, color: kText),
        ),
      ),
      home: const HomePage(),
    );
  }
}

class WhiteCard extends StatelessWidget {
  final Widget child;
  final EdgeInsets padding;
  final VoidCallback? onTap;
  const WhiteCard({super.key, required this.child, this.padding = const EdgeInsets.all(16), this.onTap});
  @override
  Widget build(BuildContext context) => Material(
        color: Colors.white,
        borderRadius: BorderRadius.circular(14),
        child: InkWell(
          borderRadius: BorderRadius.circular(14),
          onTap: onTap,
          child: Container(
            padding: padding,
            decoration: BoxDecoration(
              borderRadius: BorderRadius.circular(14),
              border: Border.all(color: kLine),
            ),
            child: child,
          ),
        ),
      );
}

// ---------- स्क्रीन 1: मुख्य पान ----------
class HomePage extends StatefulWidget {
  const HomePage({super.key});
  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  String ps = 'सर्व';

  List<Case> get cases => ps == 'सर्व' ? sampleCases : sampleCases.where((c) => c.ps == ps).toList();

  @override
  Widget build(BuildContext context) {
    final list = cases;
    final psList = ['सर्व', ...{for (final c in sampleCases) c.ps}];
    final arrestedPending = list.where((c) => c.arrested == 0).length;
    return Scaffold(
      appBar: AppBar(
        title: const Text('MCR · प्रलंबित गुन्हे'),
        actions: [
          IconButton(onPressed: () => _soon(context), icon: const Icon(Icons.sync_rounded)),
          const SizedBox(width: 4),
        ],
      ),
      body: SafeArea(
        child: Center(
          child: ConstrainedBox(
            constraints: const BoxConstraints(maxWidth: 560),
            child: ListView(
              padding: const EdgeInsets.fromLTRB(16, 12, 16, 24),
              children: [
                const Text('उपविभाग – नमुना माहिती', style: TextStyle(color: kMuted, fontSize: 13)),
                const SizedBox(height: 10),
                SizedBox(
                  height: 38,
                  child: ListView.separated(
                    scrollDirection: Axis.horizontal,
                    itemCount: psList.length,
                    separatorBuilder: (_, _) => const SizedBox(width: 8),
                    itemBuilder: (_, i) {
                      final sel = psList[i] == ps;
                      return ChoiceChip(
                        label: Text(psList[i]),
                        selected: sel,
                        showCheckmark: false,
                        selectedColor: kBlue,
                        backgroundColor: Colors.white,
                        side: BorderSide(color: sel ? kBlue : kLine),
                        labelStyle: TextStyle(color: sel ? Colors.white : kText, fontWeight: FontWeight.w600),
                        onSelected: (_) => setState(() => ps = psList[i]),
                      );
                    },
                  ),
                ),
                const SizedBox(height: 14),
                WhiteCard(
                  onTap: () => _openList(context, null),
                  child: Row(
                    children: [
                      Container(
                        width: 48,
                        height: 48,
                        decoration: BoxDecoration(color: const Color(0xFFEFF4FF), borderRadius: BorderRadius.circular(12)),
                        child: const Icon(Icons.folder_open_rounded, color: kBlue),
                      ),
                      const SizedBox(width: 14),
                      Expanded(
                        child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
                          const Text('एकूण प्रलंबित गुन्हे', style: TextStyle(color: kMuted, fontSize: 13)),
                          Text('${list.length}', style: const TextStyle(fontSize: 30, fontWeight: FontWeight.w700, color: kText, height: 1.2)),
                        ]),
                      ),
                      Column(crossAxisAlignment: CrossAxisAlignment.end, children: [
                        const Text('आरोपी अटक बाकी', style: TextStyle(color: kMuted, fontSize: 12)),
                        Text('$arrestedPending', style: const TextStyle(fontSize: 18, fontWeight: FontWeight.w700, color: kBlue)),
                      ]),
                    ],
                  ),
                ),
                const SizedBox(height: 18),
                const Text('कालावधीनुसार', style: TextStyle(fontSize: 15, fontWeight: FontWeight.w700, color: kText)),
                const SizedBox(height: 10),
                GridView(
                  gridDelegate: const SliverGridDelegateWithFixedCrossAxisCount(
                    crossAxisCount: 2,
                    mainAxisSpacing: 12,
                    crossAxisSpacing: 12,
                    mainAxisExtent: 140,
                  ),
                  shrinkWrap: true,
                  physics: const NeverScrollableScrollPhysics(),
                  children: [
                    for (final b in buckets)
                      AgingTile(
                        b: b,
                        count: list.where((c) => c.bucket.key == b.key).length,
                        total: list.length,
                        onTap: () => _openList(context, b),
                      ),
                  ],
                ),
                const SizedBox(height: 18),
                const Text('सर्वात जुने गुन्हे', style: TextStyle(fontSize: 15, fontWeight: FontWeight.w700, color: kText)),
                const SizedBox(height: 10),
                for (final c in (list.toList()..sort((a, b) => b.days.compareTo(a.days))).take(3)) ...[
                  CaseCard(c: c),
                  const SizedBox(height: 10),
                ],
                const SizedBox(height: 4),
                FilledButton(
                  style: FilledButton.styleFrom(
                    backgroundColor: kBlue,
                    minimumSize: const Size.fromHeight(50),
                    shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
                  ),
                  onPressed: () => _openList(context, null),
                  child: const Text('सर्व प्रलंबित यादी पाहा', style: TextStyle(fontSize: 15, fontWeight: FontWeight.w600)),
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }

  void _openList(BuildContext context, Bucket? b) => Navigator.of(context).push(
        MaterialPageRoute(builder: (_) => ListPage(initial: b, ps: ps)),
      );
}

void _soon(BuildContext context) => ScaffoldMessenger.of(context)
  ..hideCurrentSnackBar()
  ..showSnackBar(const SnackBar(content: Text('हा डेमो आहे – ही सुविधा पुढील आवृत्तीत'), behavior: SnackBarBehavior.floating));

class AgingTile extends StatelessWidget {
  final Bucket b;
  final int count, total;
  final VoidCallback onTap;
  const AgingTile({super.key, required this.b, required this.count, required this.total, required this.onTap});
  @override
  Widget build(BuildContext context) {
    final pct = total == 0 ? 0.0 : count / total;
    return WhiteCard(
      onTap: onTap,
      padding: const EdgeInsets.all(14),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Row(children: [
            Container(width: 10, height: 10, decoration: BoxDecoration(color: b.color, shape: BoxShape.circle)),
            const SizedBox(width: 8),
            Expanded(
              child: Text(b.title,
                  maxLines: 1, overflow: TextOverflow.ellipsis,
                  style: const TextStyle(fontSize: 13, fontWeight: FontWeight.w600, color: kText)),
            ),
          ]),
          const Spacer(),
          Text('$count', style: TextStyle(fontSize: 32, fontWeight: FontWeight.w700, color: b.color, height: 1.1)),
          Text(b.sub, style: const TextStyle(fontSize: 11.5, color: kMuted)),
          const SizedBox(height: 8),
          ClipRRect(
            borderRadius: BorderRadius.circular(4),
            child: LinearProgressIndicator(value: pct, minHeight: 5, color: b.color, backgroundColor: b.tint),
          ),
        ],
      ),
    );
  }
}

// ---------- स्क्रीन 2: यादी ----------
class ListPage extends StatefulWidget {
  final Bucket? initial;
  final String ps;
  const ListPage({super.key, this.initial, required this.ps});
  @override
  State<ListPage> createState() => _ListPageState();
}

class _ListPageState extends State<ListPage> {
  Bucket? filter;
  String q = '';

  @override
  void initState() {
    super.initState();
    filter = widget.initial;
  }

  @override
  Widget build(BuildContext context) {
    var list = sampleCases.where((c) => widget.ps == 'सर्व' || c.ps == widget.ps);
    if (filter != null) list = list.where((c) => c.bucket.key == filter!.key);
    if (q.isNotEmpty) {
      list = list.where((c) => '${c.crNo} ${c.ps} ${c.head} ${c.io} ${c.kalam}'.toLowerCase().contains(q.toLowerCase()));
    }
    final items = list.toList()..sort((a, b) => b.days.compareTo(a.days));
    return Scaffold(
      appBar: AppBar(title: Text(filter == null ? 'प्रलंबित यादी' : filter!.title)),
      body: SafeArea(
        child: Center(
          child: ConstrainedBox(
            constraints: const BoxConstraints(maxWidth: 560),
            child: Column(
              children: [
                Padding(
                  padding: const EdgeInsets.fromLTRB(16, 10, 16, 6),
                  child: TextField(
                    onChanged: (v) => setState(() => q = v),
                    decoration: InputDecoration(
                      hintText: 'गु.र.नं., शीर्षक, IO शोधा',
                      prefixIcon: const Icon(Icons.search_rounded),
                      filled: true,
                      fillColor: Colors.white,
                      isDense: true,
                      border: OutlineInputBorder(borderRadius: BorderRadius.circular(12), borderSide: const BorderSide(color: kLine)),
                      enabledBorder: OutlineInputBorder(borderRadius: BorderRadius.circular(12), borderSide: const BorderSide(color: kLine)),
                    ),
                  ),
                ),
                SizedBox(
                  height: 44,
                  child: ListView(
                    scrollDirection: Axis.horizontal,
                    padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 4),
                    children: [
                      _chip(null, 'सर्व', kBlue),
                      for (final b in buckets) _chip(b, b.title, b.color),
                    ],
                  ),
                ),
                Padding(
                  padding: const EdgeInsets.fromLTRB(16, 4, 16, 4),
                  child: Row(children: [
                    Text('${items.length} गुन्हे', style: const TextStyle(color: kMuted, fontSize: 13)),
                    const Spacer(),
                    const Text('जुने आधी', style: TextStyle(color: kMuted, fontSize: 13)),
                  ]),
                ),
                Expanded(
                  child: items.isEmpty
                      ? const Center(child: Text('काहीही सापडले नाही', style: TextStyle(color: kMuted)))
                      : ListView.separated(
                          padding: const EdgeInsets.fromLTRB(16, 4, 16, 24),
                          itemCount: items.length,
                          separatorBuilder: (_, _) => const SizedBox(height: 10),
                          itemBuilder: (_, i) => CaseCard(c: items[i]),
                        ),
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }

  Widget _chip(Bucket? b, String label, Color color) {
    final sel = filter?.key == b?.key;
    return Padding(
      padding: const EdgeInsets.only(right: 8),
      child: ChoiceChip(
        label: Text(label),
        selected: sel,
        showCheckmark: false,
        selectedColor: color,
        backgroundColor: Colors.white,
        side: BorderSide(color: sel ? color : kLine),
        labelStyle: TextStyle(color: sel ? Colors.white : kText, fontWeight: FontWeight.w600, fontSize: 13),
        onSelected: (_) => setState(() => filter = b),
      ),
    );
  }
}

class CaseCard extends StatelessWidget {
  final Case c;
  const CaseCard({super.key, required this.c});
  @override
  Widget build(BuildContext context) {
    final b = c.bucket;
    return WhiteCard(
      onTap: () => showCaseSheet(context, c),
      padding: const EdgeInsets.fromLTRB(14, 12, 14, 12),
      child: Row(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Container(width: 4, height: 64, decoration: BoxDecoration(color: b.color, borderRadius: BorderRadius.circular(2))),
          const SizedBox(width: 12),
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(children: [
                  Text('गु.र.नं. ${c.crNo}', style: const TextStyle(fontWeight: FontWeight.w700, fontSize: 15, color: kText)),
                  const SizedBox(width: 8),
                  Text('· ${c.ps}', style: const TextStyle(color: kMuted, fontSize: 13)),
                ]),
                const SizedBox(height: 2),
                Text('${c.head} · ${c.kalam}', style: const TextStyle(fontSize: 13.5, color: kText)),
                const SizedBox(height: 6),
                Row(children: [
                  const Icon(Icons.person_outline_rounded, size: 15, color: kMuted),
                  const SizedBox(width: 4),
                  Expanded(
                    child: Text(c.io, maxLines: 1, overflow: TextOverflow.ellipsis, style: const TextStyle(fontSize: 12.5, color: kMuted)),
                  ),
                ]),
              ],
            ),
          ),
          const SizedBox(width: 8),
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
            decoration: BoxDecoration(color: b.tint, borderRadius: BorderRadius.circular(8)),
            child: Column(children: [
              Text('${c.days}', style: TextStyle(color: b.color, fontWeight: FontWeight.w700, fontSize: 16, height: 1.1)),
              Text('दिवस', style: TextStyle(color: b.color, fontSize: 11)),
            ]),
          ),
        ],
      ),
    );
  }
}

// ---------- स्क्रीन 3: तपशील (bottom sheet) ----------
void showCaseSheet(BuildContext context, Case c) {
  final b = c.bucket;
  showModalBottomSheet(
    context: context,
    isScrollControlled: true,
    backgroundColor: Colors.white,
    showDragHandle: true,
    constraints: const BoxConstraints(maxWidth: 560),
    shape: const RoundedRectangleBorder(borderRadius: BorderRadius.vertical(top: Radius.circular(20))),
    builder: (ctx) => DraggableScrollableSheet(
      expand: false,
      initialChildSize: 0.72,
      maxChildSize: 0.92,
      minChildSize: 0.4,
      builder: (ctx, sc) => ListView(
        controller: sc,
        padding: const EdgeInsets.fromLTRB(20, 0, 20, 24),
        children: [
          Row(children: [
            Expanded(
              child: Text('गु.र.नं. ${c.crNo}', style: const TextStyle(fontSize: 20, fontWeight: FontWeight.w700, color: kText)),
            ),
            Container(
              padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 5),
              decoration: BoxDecoration(color: b.tint, borderRadius: BorderRadius.circular(20)),
              child: Text('${c.days} दिवस · ${b.title}', style: TextStyle(color: b.color, fontWeight: FontWeight.w600, fontSize: 12.5)),
            ),
          ]),
          const SizedBox(height: 4),
          Text('पो.स्टे. ${c.ps}', style: const TextStyle(color: kMuted)),
          const SizedBox(height: 16),
          _info('शीर्षक', c.head),
          _info('कलम', c.kalam),
          _info('दाखल दिनांक', fmtDate(c.regDate)),
          _info('फिर्यादी', c.complainant),
          _info('तपासी अधिकारी', c.io),
          _info('सद्यस्थिती', c.stage),
          const SizedBox(height: 12),
          Row(children: [
            Expanded(child: _stat('आरोपी', '${c.accused}', kText)),
            const SizedBox(width: 10),
            Expanded(child: _stat('अटक', '${c.arrested}', const Color(0xFF16A34A))),
            const SizedBox(width: 10),
            Expanded(child: _stat('अटक बाकी', '${c.accused - c.arrested}', const Color(0xFFDC2626))),
          ]),
          const SizedBox(height: 18),
          const Text('शेवटच्या नोंदी', style: TextStyle(fontWeight: FontWeight.w700, color: kText)),
          const SizedBox(height: 8),
          _note(fmtDate(c.regDate), 'गुन्हा दाखल'),
          _note(fmtDate(c.regDate.add(Duration(days: (c.days * .4).round()))), 'घटनास्थळ पंचनामा व जबाब'),
          _note(fmtDate(c.regDate.add(Duration(days: (c.days * .8).round()))), c.stage),
          const SizedBox(height: 18),
          Row(children: [
            Expanded(
              child: OutlinedButton.icon(
                style: OutlinedButton.styleFrom(
                  minimumSize: const Size.fromHeight(48),
                  foregroundColor: kBlue,
                  side: const BorderSide(color: kBlue),
                  shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
                ),
                onPressed: () => _soon(ctx),
                icon: const Icon(Icons.call_outlined),
                label: const Text('IO ला फोन'),
              ),
            ),
            const SizedBox(width: 10),
            Expanded(
              child: FilledButton.icon(
                style: FilledButton.styleFrom(
                  minimumSize: const Size.fromHeight(48),
                  backgroundColor: kBlue,
                  shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
                ),
                onPressed: () => _soon(ctx),
                icon: const Icon(Icons.edit_note_rounded),
                label: const Text('नोंद जोडा'),
              ),
            ),
          ]),
        ],
      ),
    ),
  );
}

Widget _info(String k, String v) => Container(
      padding: const EdgeInsets.symmetric(vertical: 10),
      decoration: const BoxDecoration(border: Border(bottom: BorderSide(color: kLine))),
      child: Row(crossAxisAlignment: CrossAxisAlignment.start, children: [
        SizedBox(width: 120, child: Text(k, style: const TextStyle(color: kMuted, fontSize: 13.5))),
        Expanded(child: Text(v, style: const TextStyle(color: kText, fontSize: 14.5, fontWeight: FontWeight.w600))),
      ]),
    );

Widget _stat(String k, String v, Color color) => Container(
      padding: const EdgeInsets.symmetric(vertical: 10),
      decoration: BoxDecoration(color: kBg, borderRadius: BorderRadius.circular(12)),
      child: Column(children: [
        Text(v, style: TextStyle(fontSize: 20, fontWeight: FontWeight.w700, color: color)),
        Text(k, style: const TextStyle(fontSize: 12, color: kMuted)),
      ]),
    );

Widget _note(String date, String text) => Padding(
      padding: const EdgeInsets.only(bottom: 8),
      child: Row(crossAxisAlignment: CrossAxisAlignment.start, children: [
        Container(margin: const EdgeInsets.only(top: 6), width: 8, height: 8, decoration: const BoxDecoration(color: kBlue, shape: BoxShape.circle)),
        const SizedBox(width: 10),
        SizedBox(width: 92, child: Text(date, style: const TextStyle(color: kMuted, fontSize: 13))),
        Expanded(child: Text(text, style: const TextStyle(color: kText, fontSize: 13.5))),
      ]),
    );
