from pathlib import Path
import shutil
import zipfile
from datetime import datetime
import lxml.etree as ET

ROOT = Path(".")
SORTIES = ROOT / "sorties"
TARGET = SORTIES / "Annuaire_HCP_2025_FINAL.docx"

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

PATCHES = [{'original_index': 26,
  'old_text': '2022',
  'new_text': '2024',
  'new_nodes': ['2024'],
  'before': [(-2, 'Organisation administrative'), (-4, 'الوحدة'), (-6, 'Unité'), (-8, 'أهم المؤشرات الإحصائية لعمالة مكناس')],
  'after': [(2, 'التظيم الإداري'), (3, 'Nombre de Communes dont:'), (4, 'Nombre'), (6, 'العدد')]},
 {'original_index': 60,
  'old_text': '2021-2022',
  'new_text': '2023-2024',
  'new_nodes': ['2023-2024', ''],
  'before': [(-1, 'Santé'), (-2, 'المقيمون بالوسط الحضري'), (-3, 'نسمة'), (-4, '780 290')],
  'after': [(2, 'الصحة'), (3, 'Médecins ( public + privé )'), (4, 'Nombre'), (6, 'العدد')]},
 {'original_index': 65,
  'old_text': '724',
  'new_text': '804',
  'new_nodes': ['804'],
  'before': [(-1, 'Nombre'), (-2, 'Médecins ( public + privé )'), (-3, 'الصحة'), (-6, 'Santé')],
  'after': [(1, 'العدد'), (2, 'الأطباء ) العمومي + الخاص ('), (3, 'Dont secteur public'), (4, 'Nombre')]},
 {'original_index': 70,
  'old_text': '264',
  'new_text': '261',
  'new_nodes': ['261'],
  'before': [(-1, 'Nombre'), (-2, 'Dont secteur public'), (-3, 'الأطباء ) العمومي + الخاص ('), (-4, 'العدد')],
  'after': [(1, 'العدد'), (2, 'منهم : بالقطاع العمومي'), (3, 'Lits disponibles pour l’hospitalisation des malades'), (4, 'Nombre')]},
 {'original_index': 75,
  'old_text': '638',
  'new_text': '610',
  'new_nodes': ['610'],
  'before': [(-1, 'Nombre'), (-2, 'Lits disponibles pour l’hospitalisation des malades'), (-3, 'منهم : بالقطاع العمومي'), (-4, 'العدد')],
  'after': [(1, 'العدد'), (2, 'الأسرة المتواجدة لستشفاء المرض ى'), (3, 'Admissions en hospitalisation complète'), (4, 'Nombre')]},
 {'original_index': 80,
  'old_text': '23 915',
  'new_text': '32 466',
  'new_nodes': ['32 466', '', ''],
  'before': [(-1, 'Nombre'), (-2, 'Admissions en hospitalisation complète'), (-3, 'الأسرة المتواجدة لستشفاء المرض ى'), (-4, 'العدد')],
  'after': [(1, 'العدد'), (2, 'الدخول للإستشفاء الشامل'), (3, "Journées d'hospitalisation"), (4, 'Nombre')]},
 {'original_index': 85,
  'old_text': '114 825',
  'new_text': '131 058',
  'new_nodes': ['131 058', '', ''],
  'before': [(-1, 'Nombre'), (-2, "Journées d'hospitalisation"), (-3, 'الدخول للإستشفاء الشامل'), (-4, 'العدد')],
  'after': [(1, 'العدد'), (2, 'أيـــام الإستشفاء'), (3, "Taux moyen d'occupation"), (7, 'متوسط معدل الإيواء')]},
 {'original_index': 90,
  'old_text': '49,5',
  'new_text': '61,3',
  'new_nodes': ['61,3'],
  'before': [(-2, "Taux moyen d'occupation"), (-3, 'أيـــام الإستشفاء'), (-4, 'العدد'), (-6, 'Nombre')],
  'after': [(2, 'متوسط معدل الإيواء'), (3, 'Durée moyenne de séjour'), (4, 'Jour'), (6, 'اليوم')]},
 {'original_index': 95,
  'old_text': '4,8',
  'new_text': '4',
  'new_nodes': ['4'],
  'before': [(-1, 'Jour'), (-2, 'Durée moyenne de séjour'), (-3, 'متوسط معدل الإيواء'), (-7, "Taux moyen d'occupation")],
  'after': [(1, 'اليوم'), (2, 'متوسط مدة الإقامة'), (3, 'Taux de rotation des lits'), (7, 'معدل التناوب على الأسرة')]},
 {'original_index': 100,
  'old_text': '37,5',
  'new_text': '53,2',
  'new_nodes': ['53,2'],
  'before': [(-2, 'Taux de rotation des lits'), (-3, 'متوسط مدة الإقامة'), (-4, 'اليوم'), (-6, 'Jour')],
  'after': [(2, 'معدل التناوب على الأسرة'), (3, 'Education'), (6, 'التعليم'), (7, 'Effectifs inscrits dans les écoles coraniques')]},
 {'original_index': 104,
  'old_text': '2022-2023',
  'new_text': '2024-2025',
  'new_nodes': ['2024-2025', ''],
  'before': [(-1, 'Education'), (-2, 'معدل التناوب على الأسرة'), (-6, 'Taux de rotation des lits'), (-7, 'متوسط مدة الإقامة')],
  'after': [(2, 'التعليم'), (3, 'Effectifs inscrits dans les écoles coraniques'), (4, 'Nombre'), (6, 'العدد')]},
 {'original_index': 109,
  'old_text': '4 486',
  'new_text': '4 236',
  'new_nodes': ['4 236', '', ''],
  'before': [(-1, 'Nombre'), (-2, 'Effectifs inscrits dans les écoles coraniques'), (-3, 'التعليم'), (-6, 'Education')],
  'after': [(1, 'العدد'), (2, 'المسجلون بالكتاتيب القرآنية'), (3, 'Effectifs inscrits dans le préscolaire moderne'), (4, 'Nombre')]},
 {'original_index': 114,
  'old_text': '15 383',
  'new_text': '17 660',
  'new_nodes': ['17 660', '', ''],
  'before': [(-1, 'Nombre'), (-2, 'Effectifs inscrits dans le préscolaire moderne'), (-3, 'المسجلون بالكتاتيب القرآنية'), (-4, 'العدد')],
  'after': [(1, 'العدد'), (2, 'المسجلون بالتعليم الأولي العصري'), (3, "Effectifs inscrits dans l'enseignement primaire public"), (4, 'Nombre')]},
 {'original_index': 119,
  'old_text': '84 512',
  'new_text': '79 847',
  'new_nodes': ['79 847', '', ''],
  'before': [(-1, 'Nombre'), (-2, "Effectifs inscrits dans l'enseignement primaire public"), (-3, 'المسجلون بالتعليم الأولي العصري'), (-4, 'العدد')],
  'after': [(1, 'العدد'), (2, 'المسجلون بالتعليم الإبتدائي العمومي'), (3, "Effectifs inscrits dans l'enseignement secondaire"), (4, 'Collégial public')]},
 {'original_index': 127,
  'old_text': '43 917',
  'new_text': '47 056',
  'new_nodes': ['47 056', '', ''],
  'before': [(-2, 'Nombre'), (-4, 'Collégial public'), (-5, "Effectifs inscrits dans l'enseignement secondaire"), (-6, 'المسجلون بالتعليم الإبتدائي العمومي')],
  'after': [(2, 'العدد'),
            (4, 'المسجلون بالتعليم الثانوي الإعدادي العمومي'),
            (5, "Effectifs inscrits dans l'enseignement secondaire"),
            (6, 'qualifiant public')]},
 {'original_index': 137,
  'old_text': '26 658',
  'new_text': '29 531',
  'new_nodes': ['29 531', '', ''],
  'before': [(-2, 'Nombre'),
             (-4, 'qualifiant public'),
             (-5, "Effectifs inscrits dans l'enseignement secondaire"),
             (-6, 'المسجلون بالتعليم الثانوي الإعدادي العمومي')],
  'after': [(2, 'العدد'), (4, 'المسجلون بالتعليم الثانوي التأهيلي العمومي'), (5, "Effectifs inscrits dans l'enseignement supérieur"), (6, 'Nombre')]},
 {'original_index': 144,
  'old_text': '79 947',
  'new_text': '82 196',
  'new_nodes': ['82 196', '', ''],
  'before': [(-1, 'Nombre'), (-2, "Effectifs inscrits dans l'enseignement supérieur"), (-3, 'المسجلون بالتعليم الثانوي التأهيلي العمومي'), (-5, 'العدد')],
  'after': [(1, 'العدد'), (2, 'المسجلون بالتعليم العالي'), (3, "Enseignants de l'enseignement primaire public"), (4, 'Nombre')]},
 {'original_index': 149,
  'old_text': '2 889',
  'new_text': '2 910',
  'new_nodes': ['2 910', '', ''],
  'before': [(-1, 'Nombre'), (-2, "Enseignants de l'enseignement primaire public"), (-3, 'المسجلون بالتعليم العالي'), (-4, 'العدد')],
  'after': [(1, 'العدد'), (2, 'أساتذة التعليم الإبتدائي العمومي'), (3, "Enseignants de l'enseignement secondaire Collégial"), (4, 'public')]},
 {'original_index': 157,
  'old_text': '1 835',
  'new_text': '1 940',
  'new_nodes': ['1 940', '', ''],
  'before': [(-2, 'Nombre'), (-4, 'public'), (-5, "Enseignants de l'enseignement secondaire Collégial"), (-6, 'أساتذة التعليم الإبتدائي العمومي')],
  'after': [(2, 'العدد'), (3, 'أساتذة التعليم الثانوي الإعدادي العمومي'), (4, 'Enseignants dans le secondaire qualifiant public'), (5, 'Nombre')]},
 {'original_index': 163,
  'old_text': '1 609',
  'new_text': '1 696',
  'new_nodes': ['1 696', '', ''],
  'before': [(-1, 'Nombre'), (-2, 'Enseignants dans le secondaire qualifiant public'), (-3, 'أساتذة التعليم الثانوي الإعدادي العمومي'), (-4, 'العدد')],
  'after': [(1, 'العدد'), (2, 'أساتدة التعليم الثانوي التأهيلي العمومي'), (3, 'Enseignants dans les universités ( permanents)'), (4, 'Nombre')]},
 {'original_index': 168,
  'old_text': '1 123',
  'new_text': '1 214',
  'new_nodes': ['1 214', '', ''],
  'before': [(-1, 'Nombre'), (-2, 'Enseignants dans les universités ( permanents)'), (-3, 'أساتدة التعليم الثانوي التأهيلي العمومي'), (-4, 'العدد')],
  'after': [(1, 'العدد'), (2, 'أساتدة التعليم الجامعي ) الهيئة االدائمة ('), (3, 'Emploi / Salaire / Prix'), (4, '2022')]},
 {'original_index': 221,
  'old_text': '2021-2022',
  'new_text': '2023-2024',
  'new_nodes': ['2023-2024', ''],
  'before': [(-2, 'Agriculture'), (-4, 'الوحدة'), (-6, 'Unité'), (-8, 'أهم المؤشرات الإحصائية لعمالة مكناس')],
  'after': [(2, 'الفلاحة'), (3, 'Superficie des céréales'), (4, 'Milliers Ha'), (6, 'ألف هكتار')]},
 {'original_index': 226,
  'old_text': '64,1',
  'new_text': '61,2',
  'new_nodes': ['61,2'],
  'before': [(-1, 'Milliers Ha'), (-2, 'Superficie des céréales'), (-3, 'الفلاحة'), (-7, 'Agriculture')],
  'after': [(1, 'ألف هكتار'), (2, 'مساحة الحبوب'), (3, 'Production des céréales'), (4, 'Milliers Qt')]},
 {'original_index': 231,
  'old_text': '1067,3',
  'new_text': '1 590,3',
  'new_nodes': ['1 590,3'],
  'before': [(-1, 'Milliers Qt'), (-2, 'Production des céréales'), (-3, 'مساحة الحبوب'), (-4, 'ألف هكتار')],
  'after': [(1, 'ألف قنطار'), (2, 'إنتاج الحبوب'), (3, 'Superficie des légumineuses'), (4, 'Milliers Ha')]},
 {'original_index': 236,
  'old_text': '10,7',
  'new_text': '7,4',
  'new_nodes': ['7,4'],
  'before': [(-1, 'Milliers Ha'), (-2, 'Superficie des légumineuses'), (-3, 'إنتاج الحبوب'), (-4, 'ألف قنطار')],
  'after': [(1, 'ألف هكتار'), (2, 'مساحة القطاني'), (3, 'Production des légumineuses'), (4, 'Milliers Qt')]},
 {'original_index': 241,
  'old_text': '107,2',
  'new_text': '72,3',
  'new_nodes': ['72,3'],
  'before': [(-1, 'Milliers Qt'), (-2, 'Production des légumineuses'), (-3, 'مساحة القطاني'), (-4, 'ألف هكتار')],
  'after': [(1, 'ألف قنطار'), (2, 'إنتاج القطاني'), (3, 'Cheptel : Bovins'), (4, 'Milliers de têtes')]},
 {'original_index': 246,
  'old_text': '44,9',
  'new_text': '34,9',
  'new_nodes': ['34,9'],
  'before': [(-1, 'Milliers de têtes'), (-2, 'Cheptel : Bovins'), (-3, 'إنتاج القطاني'), (-4, 'ألف قنطار')],
  'after': [(1, 'آلف الرؤوس'), (2, 'الماشية : البقر'), (3, 'Ovins'), (4, 'Milliers de têtes')]},
 {'original_index': 251,
  'old_text': '235,1',
  'new_text': '156,7',
  'new_nodes': ['156,7'],
  'before': [(-1, 'Milliers de têtes'), (-2, 'Ovins'), (-3, 'الماشية : البقر'), (-4, 'آلف الرؤوس')],
  'after': [(1, 'آلف الرؤوس'), (2, 'الغنم'), (3, 'Caprins'), (4, 'Milliers de têtes')]},
 {'original_index': 256,
  'old_text': '10,2',
  'new_text': '7,0',
  'new_nodes': ['7,0'],
  'before': [(-1, 'Milliers de têtes'), (-2, 'Caprins'), (-3, 'الغنم'), (-4, 'آلف الرؤوس')],
  'after': [(1, 'آلف الرؤوس'), (2, 'الماعز'), (3, 'Energie et Eau'), (7, 'الطاقة والماء')]},
 {'original_index': 261,
  'old_text': '2022',
  'new_text': '2023-2024',
  'new_nodes': ['2023-2024'],
  'before': [(-2, 'Energie et Eau'), (-3, 'الماعز'), (-4, 'آلف الرؤوس'), (-6, 'Milliers de têtes')],
  'after': [(2, 'الطاقة والماء'), (3, "Ventes d'énergie électrique"), (4, 'En millions de Kwh'), (6, 'مليون كيلواط/ ساعة')]},
 {'original_index': 266,
  'old_text': '770,9',
  'new_text': '783,1',
  'new_nodes': ['783,1'],
  'before': [(-1, 'En millions de Kwh'), (-2, "Ventes d'énergie électrique"), (-3, 'الطاقة والماء'), (-7, 'Energie et Eau')],
  'after': [(1, 'مليون كيلواط/ ساعة'), (2, 'مبيعات الطاقة الكهربائية'), (3, "Production de l'eau potable"), (4, 'En milliers de mètre cube')]},
 {'original_index': 272,
  'old_text': '3 086',
  'new_text': '3 071',
  'new_nodes': ['3 071', '', ''],
  'before': [(-1, '(m3)'), (-2, 'En milliers de mètre cube'), (-3, "Production de l'eau potable"), (-4, 'مبيعات الطاقة الكهربائية')],
  'after': [(1, 'ألف م3'), (2, 'إنتاج الماء الصالح للشرب'), (3, 'Tourisme'), (7, 'السياحة')]},
 {'original_index': 277,
  'old_text': '2022',
  'new_text': '2022-2024',
  'new_nodes': ['2022-2024'],
  'before': [(-2, 'Tourisme'), (-3, 'إنتاج الماء الصالح للشرب'), (-4, 'ألف م3'), (-6, '(m3)')],
  'after': [(2, 'السياحة'), (3, "Nombre d'établissement classés"), (4, 'Unité'), (6, 'الوحدة')]},
 {'original_index': 282,
  'old_text': '82',
  'new_text': '81',
  'new_nodes': ['81'],
  'before': [(-1, 'Unité'), (-2, "Nombre d'établissement classés"), (-3, 'السياحة'), (-7, 'Tourisme')],
  'after': [(1, 'الوحدة'), (2, 'عدد المؤسسات المصنفة'), (3, 'Nombre de lits disponibles'), (4, '(établissement classés)')]},
 {'original_index': 288,
  'old_text': '4 425',
  'new_text': '4 347',
  'new_nodes': ['4 347', '', ''],
  'before': [(-1, 'Lits'), (-2, '(établissement classés)'), (-3, 'Nombre de lits disponibles'), (-4, 'عدد المؤسسات المصنفة')],
  'after': [(1, 'الأسرة'), (2, 'عدد الأسرة )المؤسسات المصنفة ('), (3, 'Nuitées touristique (établissement classés )'), (4, 'Nuitées')]},
 {'original_index': 298,
  'old_text': '106 231',
  'new_text': '122 119',
  'new_nodes': ['122 119', '', ''],
  'before': [(-1, 'Nombre'), (-2, 'Arrivées (établissement classés )'), (-3, 'الليالي السياحية المنجزة بالمؤسسات المصنفة'), (-4, 'الليالي السياحية')],
  'after': [(1, 'العدد'), (2, 'القادمين للمؤسسات المصنفة'), (3, 'Habitat et Urbanisation'), (5, '2020')]},
 {'original_index': 333,
  'old_text': '2022',
  'new_text': '2024',
  'new_nodes': ['2024'],
  'before': [(-2, 'Transports'), (-3, 'المساحة المسقفة'), (-4, 'بالمترالمربع'), (-5, '583 906')],
  'after': [(2, 'النقل والمواصلات'), (3, 'Immatriculation des véhicules'), (4, 'Nombre'), (6, 'العدد')]},
 {'original_index': 338,
  'old_text': '4 788',
  'new_text': '4 849',
  'new_nodes': ['4 849', '', ''],
  'before': [(-1, 'Nombre'), (-2, 'Immatriculation des véhicules'), (-3, 'النقل والمواصلات'), (-7, 'Transports')],
  'after': [(1, 'العدد'), (2, 'تسجيل ا العربات'), (3, 'Permis de conduire délivrés'), (4, 'Nombre')]},
 {'original_index': 343,
  'old_text': '17 512',
  'new_text': '16 392',
  'new_nodes': ['16 392', '', ''],
  'before': [(-1, 'Nombre'), (-2, 'Permis de conduire délivrés'), (-3, 'تسجيل ا العربات'), (-4, 'العدد')],
  'after': [(1, 'العدد'),
            (2, 'رخص السياقة المسلمة'),
            (3, 'Indice des prix à la consommation ( Ville de Meknès )'),
            (5, 'الرقم الإستدلالي للأثمان عند الإستهلاك )مدينة مكناس(')]},
 {'original_index': 347,
  'old_text': '2022',
  'new_text': '2024',
  'new_nodes': ['2024'],
  'before': [(-1, 'Indice des prix à la consommation ( Ville de Meknès )'), (-2, 'رخص السياقة المسلمة'), (-3, 'العدد'), (-5, 'Nombre')],
  'after': [(1, 'الرقم الإستدلالي للأثمان عند الإستهلاك )مدينة مكناس('),
            (2, 'Produits alimentaires et boissons non alcoolisées'),
            (4, 'المواد الغذائية والمشروبات غيرالكحولية'),
            (5, 'Boissons alcoolisées et tabac')]},
 {'original_index': 350,
  'old_text': '113,4',
  'new_text': '125,7',
  'new_nodes': ['125,7'],
  'before': [(-1, 'Produits alimentaires et boissons non alcoolisées'),
             (-2, 'الرقم الإستدلالي للأثمان عند الإستهلاك )مدينة مكناس('),
             (-4, 'Indice des prix à la consommation ( Ville de Meknès )'),
             (-5, 'رخص السياقة المسلمة')],
  'after': [(1, 'المواد الغذائية والمشروبات غيرالكحولية'),
            (2, 'Boissons alcoolisées et tabac'),
            (6, 'المشروبات الكحولية والتبغ'),
            (7, "Articles d'habillement et chaussures")]},
 {'original_index': 354,
  'old_text': '130,4',
  'new_text': '140,4',
  'new_nodes': ['140,4'],
  'before': [(-2, 'Boissons alcoolisées et tabac'),
             (-3, 'المواد الغذائية والمشروبات غيرالكحولية'),
             (-5, 'Produits alimentaires et boissons non alcoolisées'),
             (-6, 'الرقم الإستدلالي للأثمان عند الإستهلاك )مدينة مكناس(')],
  'after': [(2, 'المشروبات الكحولية والتبغ'),
            (3, "Articles d'habillement et chaussures"),
            (7, 'الملابس والأحذية'),
            (8, 'Logement, eau, électricité, gaz et autres combustibles')]},
 {'original_index': 359,
  'old_text': '107,6',
  'new_text': '112,3',
  'new_nodes': ['112,3'],
  'before': [(-2, "Articles d'habillement et chaussures"),
             (-3, 'المشروبات الكحولية والتبغ'),
             (-7, 'Boissons alcoolisées et tabac'),
             (-8, 'المواد الغذائية والمشروبات غيرالكحولية')],
  'after': [(2, 'الملابس والأحذية'),
            (3, 'Logement, eau, électricité, gaz et autres combustibles'),
            (5, 'السكن والماء والكهرباء والغازومحروقات أخرى'),
            (6, 'Meubles, articles de ménage et entretien courant du foyer')]},
 {'original_index': 363,
  'old_text': '103,9',
  'new_text': '106,9',
  'new_nodes': ['106,9'],
  'before': [(-1, 'Logement, eau, électricité, gaz et autres combustibles'),
             (-2, 'الملابس والأحذية'),
             (-6, "Articles d'habillement et chaussures"),
             (-7, 'المشروبات الكحولية والتبغ')],
  'after': [(1, 'السكن والماء والكهرباء والغازومحروقات أخرى'),
            (2, 'Meubles, articles de ménage et entretien courant du foyer'),
            (4, 'الأثاث والأدوات المنزلية والصيانة العادية للمنزل'),
            (5, 'Santé')]},
 {'original_index': 366,
  'old_text': '103,9',
  'new_text': '110,7',
  'new_nodes': ['110,7'],
  'before': [(-1, 'Meubles, articles de ménage et entretien courant du foyer'),
             (-2, 'السكن والماء والكهرباء والغازومحروقات أخرى'),
             (-4, 'Logement, eau, électricité, gaz et autres combustibles'),
             (-5, 'الملابس والأحذية')],
  'after': [(1, 'الأثاث والأدوات المنزلية والصيانة العادية للمنزل'), (2, 'Santé'), (6, 'الصحة'), (7, 'Transport')]},
 {'original_index': 370,
  'old_text': '102,0',
  'new_text': '102,6',
  'new_nodes': ['102,6'],
  'before': [(-2, 'Santé'),
             (-3, 'الأثاث والأدوات المنزلية والصيانة العادية للمنزل'),
             (-5, 'Meubles, articles de ménage et entretien courant du foyer'),
             (-6, 'السكن والماء والكهرباء والغازومحروقات أخرى')],
  'after': [(2, 'الصحة'), (3, 'Transport'), (7, 'النقل'), (8, 'Communications')]},
 {'original_index': 375,
  'old_text': '120,8',
  'new_text': '121,8',
  'new_nodes': ['121,8'],
  'before': [(-2, 'Transport'), (-3, 'الصحة'), (-7, 'Santé'), (-8, 'الأثاث والأدوات المنزلية والصيانة العادية للمنزل')],
  'after': [(2, 'النقل'), (3, 'Communications'), (7, 'المواصلات'), (8, 'Loisirs et culture')]},
 {'original_index': 380,
  'old_text': '104,4',
  'new_text': '104,1',
  'new_nodes': ['104,1'],
  'before': [(-2, 'Communications'), (-3, 'النقل'), (-7, 'Transport'), (-8, 'الصحة')],
  'after': [(2, 'المواصلات'), (3, 'Loisirs et culture'), (7, 'الترفيه والثقافة'), (8, 'Enseignement')]},
 {'original_index': 385,
  'old_text': '107,7',
  'new_text': '112,9',
  'new_nodes': ['112,9'],
  'before': [(-2, 'Loisirs et culture'), (-3, 'المواصلات'), (-7, 'Communications'), (-8, 'النقل')],
  'after': [(2, 'الترفيه والثقافة'), (3, 'Enseignement'), (7, 'التعليم'), (8, 'Restaurants et hôtels')]},
 {'original_index': 390,
  'old_text': '109,3',
  'new_text': '112,3',
  'new_nodes': ['112,3'],
  'before': [(-2, 'Enseignement'), (-3, 'الترفيه والثقافة'), (-7, 'Loisirs et culture'), (-8, 'المواصلات')],
  'after': [(2, 'التعليم'), (3, 'Restaurants et hôtels'), (7, 'مطاعم وفنادق'), (8, 'Indice général')]},
 {'original_index': 395,
  'old_text': '106,2',
  'new_text': '116,1',
  'new_nodes': ['116,1'],
  'before': [(-2, 'Restaurants et hôtels'), (-3, 'التعليم'), (-7, 'Enseignement'), (-8, 'الترفيه والثقافة')],
  'after': [(2, 'مطاعم وفنادق'), (3, 'Indice général'), (7, 'الرقم الإستدلالي العام'), (8, 'Variation (en %)')]},
 {'original_index': 400,
  'old_text': '110,6',
  'new_text': '117,0',
  'new_nodes': ['117,0'],
  'before': [(-2, 'Indice général'), (-3, 'مطاعم وفنادق'), (-7, 'Restaurants et hôtels'), (-8, 'التعليم')],
  'after': [(2, 'الرقم الإستدلالي العام'), (3, 'Variation (en %)'), (7, 'نسبة التغـيـر(%)'), (9, 'POPULATION')]},
 {'original_index': 405,
  'old_text': '6,8',
  'new_text': '0,0',
  'new_nodes': ['0,0'],
  'before': [(-2, 'Variation (en %)'), (-3, 'الرقم الإستدلالي العام'), (-7, 'Indice général'), (-8, 'مطاعم وفنادق')],
  'after': [(2, 'نسبة التغـيـر(%)'), (4, 'POPULATION'), (5, 'السكان'), (8, 'Liste des Tableaux')]},
 {'original_index': 8127,
  'old_text': 'Fémini3n',
  'new_text': 'Féminin',
  'new_nodes': ['Féminin'],
  'before': [(-1, 'Total'), (-3, 'Etrangers'), (-4, 'الإنـــاث'), (-5, 'المجموع')],
  'after': [(6,
             'Cité Universitaire (Meknès)2 6681 90861الحي الجامعي (مكناس)Nation60 91538 4911 487الوطنCité Universitaire (Meknès)2 6681 90861الحي الجامعي '
             '(مكناس)Nation60 91538 4911 487الوطنSource : Annuaire Statistique du Maroc 2025'),
            (7, 'Cité Universitaire (Meknès)'),
            (8, '2 668'),
            (9, '1 908')]},
 {'original_index': 8491,
  'old_text': '20 Installations sportives',
  'new_text': '21 Installations sportives',
  'new_nodes': ['21', ' ', 'Installations', ' ', 'sportives'],
  'before': [(-6, 'النشطة الثقافية والترفيه'),
             (-7, 'ACTIVITES CULTURELLES ET LOISIRS'),
             (-9, 'المصدر: النشرة الإحصائية السنوية للمغرب 2025'),
             (-13, 'Source : Annuaire Statistique du Maroc 2025')],
  'after': [(6, '2 Salles de cinéma'),
            (7, '1 015 places'),
            (8, '2 قاعات السينما 015 1 عدد المقاعد'),
            (19, '14 maisons des jeunes 28 207 participants aux Activités normales et de rayonnement')]},
 {'original_index': 8495,
  'old_text': '20 تجهيز رياضي',
  'new_text': '21 تجهيز رياضي',
  'new_nodes': ['21', ' ', 'تجهيز', ' ', 'رياضي'],
  'before': [(-10, 'النشطة الثقافية والترفيه'),
             (-11, 'ACTIVITES CULTURELLES ET LOISIRS'),
             (-13, 'المصدر: النشرة الإحصائية السنوية للمغرب 2025'),
             (-17, 'Source : Annuaire Statistique du Maroc 2025')],
  'after': [(2, '2 Salles de cinéma'),
            (3, '1 015 places'),
            (4, '2 قاعات السينما 015 1 عدد المقاعد'),
            (15, '14 maisons des jeunes 28 207 participants aux Activités normales et de rayonnement')]},
 {'original_index': 8501,
  'old_text': '3 bibliothéques',
  'new_text': '5 bibliothéques',
  'new_nodes': ['5', ' ', 'bibliothéques'],
  'before': [(-2, '2 قاعات السينما 015 1 عدد المقاعد'), (-3, '1 015 places'), (-4, '2 Salles de cinéma'), (-16, 'النشطة الثقافية والترفيه')],
  'after': [(9, '14 maisons des jeunes 28 207 participants aux Activités normales et de rayonnement'),
            (15, '14 دار الشباب'),
            (16, '207 28 مشارك في الأنشطة العادية و الاشعاعية'),
            (18, '1 conservatoire')]},
 {'original_index': 8502,
  'old_text': '1 861 inscrits',
  'new_text': '2 308 inscrits',
  'new_nodes': ['2', ' ', '308 ', 'inscrits'],
  'before': [(-3, '2 قاعات السينما 015 1 عدد المقاعد'), (-4, '1 015 places'), (-5, '2 Salles de cinéma'), (-17, 'النشطة الثقافية والترفيه')],
  'after': [(8, '14 maisons des jeunes 28 207 participants aux Activités normales et de rayonnement'),
            (14, '14 دار الشباب'),
            (15, '207 28 مشارك في الأنشطة العادية و الاشعاعية'),
            (17, '1 conservatoire')]},
 {'original_index': 8507,
  'old_text': '3 مكتبات',
  'new_text': '5 مكتبات',
  'new_nodes': ['5', ' ', 'مكتبات'],
  'before': [(-8, '2 قاعات السينما 015 1 عدد المقاعد'), (-9, '1 015 places'), (-10, '2 Salles de cinéma'), (-22, 'النشطة الثقافية والترفيه')],
  'after': [(3, '14 maisons des jeunes 28 207 participants aux Activités normales et de rayonnement'),
            (9, '14 دار الشباب'),
            (10, '207 28 مشارك في الأنشطة العادية و الاشعاعية'),
            (12, '1 conservatoire')]},
 {'original_index': 8508,
  'old_text': '861 1 مسجل',
  'new_text': '308 2 مسجل',
  'new_nodes': ['308', ' ', '2', ' ', 'مسجل'],
  'before': [(-9, '2 قاعات السينما 015 1 عدد المقاعد'), (-10, '1 015 places'), (-11, '2 Salles de cinéma'), (-23, 'النشطة الثقافية والترفيه')],
  'after': [(2, '14 maisons des jeunes 28 207 participants aux Activités normales et de rayonnement'),
            (8, '14 دار الشباب'),
            (9, '207 28 مشارك في الأنشطة العادية و الاشعاعية'),
            (11, '1 conservatoire')]},
 {'original_index': 8520,
  'old_text': '340 élèves inscrits',
  'new_text': '339 élèves inscrits',
  'new_nodes': ['339', ' ', 'élèves', ' ', 'inscrits'],
  'before': [(-1, '1 conservatoire'),
             (-3, '207 28 مشارك في الأنشطة العادية و الاشعاعية'),
             (-4, '14 دار الشباب'),
             (-10, '14 maisons des jeunes 28 207 participants aux Activités normales et de rayonnement')],
  'after': []},
 {'original_index': 8525,
  'old_text': '1 معهد موسيقي 340 تلميذ مسجل',
  'new_text': '1 معهد موسيقي 339 تلميذ مسجل',
  'new_nodes': ['1', ' ', 'معهد موسيقي', ' ', '339', ' ', 'تلميذ', ' ', 'مسجل'],
  'before': [(-6, '1 conservatoire'),
             (-8, '207 28 مشارك في الأنشطة العادية و الاشعاعية'),
             (-9, '14 دار الشباب'),
             (-15, '14 maisons des jeunes 28 207 participants aux Activités normales et de rayonnement')],
  'after': []},
 {'original_index': 10704,
  'old_text': '3 086 milliers de mètre cube (m3) : production de la préfecture de Meknès en eau potable3 086 milliers de mètre cube (m3) : production de la '
              'préfecture de Meknès en eau potable086 3 ألف متر مكعب : إنتاج عمالة مكناس من الماء الصالح للشرب086 3 ألف متر مكعب : إنتاج عمالة مكناس من الماء '
              'الصالح للشرب',
  'new_text': '3 071 milliers de mètre cube (m3) : production de la préfecture de Meknès en eau potable3 071 milliers de mètre cube (m3) : production de la '
              'préfecture de Meknès en eau potable071 3 ألف متر مكعب : إنتاج عمالة مكناس من الماء الصالح للشرب071 3 ألف متر مكعب : إنتاج عمالة مكناس من الماء '
              'الصالح للشرب',
  'new_nodes': ['3 071 milliers de mètre cube (m',
                '3',
                ') : production de la préfecture',
                ' ',
                'de',
                ' ',
                'Meknès',
                ' ',
                'en',
                ' ',
                'eau ',
                'potable',
                '3 071 milliers de mètre cube (m',
                '3',
                ') : production de la préfecture',
                ' ',
                'de',
                ' ',
                'Meknès',
                ' ',
                'en',
                ' ',
                'eau ',
                'potable',
                '071',
                ' ',
                '3',
                ' ',
                'ألف متر مكعب ',
                ': ',
                'إنتاج',
                ' ',
                'عمالة مكناس من الماء الصالح ',
                'للشرب',
                '071',
                ' ',
                '3',
                ' ',
                'ألف متر مكعب ',
                ': ',
                'إنتاج',
                ' ',
                'عمالة مكناس من الماء الصالح ',
                'للشرب'],
  'before': [(-7, '(1) y compris les abonnés directs,les associations et les gros clientsالزبناء كبار و الجمعيات المباشرين، المشتركين ذلك في بما (1)'),
             (-10, 'المصدر: النشرة الإحصائية السنوية للمغرب2025'),
             (-11, 'Source : Annuaire Statistique du Maroc 2025'),
             (-13, 'عمالة : مكناس')],
  'after': [(6, 'ENERGIE'), (7, 'الطاقة'), (9, "11-2 Ventes d'électricité"), (10, '2-11 مبيعات الطاقة الكهربائية')]},
 {'original_index': 10705,
  'old_text': '3 086 milliers de mètre cube (m3) : production de la préfecture de Meknès en eau potable',
  'new_text': '3 071 milliers de mètre cube (m3) : production de la préfecture de Meknès en eau potable',
  'new_nodes': ['3 071 milliers de mètre cube (m', '3', ') : production de la préfecture', ' ', 'de', ' ', 'Meknès', ' ', 'en', ' ', 'eau ', 'potable'],
  'before': [(-8, '(1) y compris les abonnés directs,les associations et les gros clientsالزبناء كبار و الجمعيات المباشرين، المشتركين ذلك في بما (1)'),
             (-11, 'المصدر: النشرة الإحصائية السنوية للمغرب2025'),
             (-12, 'Source : Annuaire Statistique du Maroc 2025'),
             (-14, 'عمالة : مكناس')],
  'after': [(5, 'ENERGIE'), (6, 'الطاقة'), (8, "11-2 Ventes d'électricité"), (9, '2-11 مبيعات الطاقة الكهربائية')]},
 {'original_index': 10706,
  'old_text': '3 086 milliers de mètre cube (m3) : production de la préfecture de Meknès en eau potable',
  'new_text': '3 071 milliers de mètre cube (m3) : production de la préfecture de Meknès en eau potable',
  'new_nodes': ['3 071 milliers de mètre cube (m', '3', ') : production de la préfecture', ' ', 'de', ' ', 'Meknès', ' ', 'en', ' ', 'eau ', 'potable'],
  'before': [(-9, '(1) y compris les abonnés directs,les associations et les gros clientsالزبناء كبار و الجمعيات المباشرين، المشتركين ذلك في بما (1)'),
             (-12, 'المصدر: النشرة الإحصائية السنوية للمغرب2025'),
             (-13, 'Source : Annuaire Statistique du Maroc 2025'),
             (-15, 'عمالة : مكناس')],
  'after': [(4, 'ENERGIE'), (5, 'الطاقة'), (7, "11-2 Ventes d'électricité"), (8, '2-11 مبيعات الطاقة الكهربائية')]},
 {'original_index': 10707,
  'old_text': '086 3 ألف متر مكعب : إنتاج عمالة مكناس من الماء الصالح للشرب',
  'new_text': '071 3 ألف متر مكعب : إنتاج عمالة مكناس من الماء الصالح للشرب',
  'new_nodes': ['071', ' ', '3', ' ', 'ألف متر مكعب ', ': ', 'إنتاج', ' ', 'عمالة مكناس من الماء الصالح ', 'للشرب'],
  'before': [(-10, '(1) y compris les abonnés directs,les associations et les gros clientsالزبناء كبار و الجمعيات المباشرين، المشتركين ذلك في بما (1)'),
             (-13, 'المصدر: النشرة الإحصائية السنوية للمغرب2025'),
             (-14, 'Source : Annuaire Statistique du Maroc 2025'),
             (-16, 'عمالة : مكناس')],
  'after': [(3, 'ENERGIE'), (4, 'الطاقة'), (6, "11-2 Ventes d'électricité"), (7, '2-11 مبيعات الطاقة الكهربائية')]},
 {'original_index': 10708,
  'old_text': '086 3 ألف متر مكعب : إنتاج عمالة مكناس من الماء الصالح للشرب',
  'new_text': '071 3 ألف متر مكعب : إنتاج عمالة مكناس من الماء الصالح للشرب',
  'new_nodes': ['071', ' ', '3', ' ', 'ألف متر مكعب ', ': ', 'إنتاج', ' ', 'عمالة مكناس من الماء الصالح ', 'للشرب'],
  'before': [(-11, '(1) y compris les abonnés directs,les associations et les gros clientsالزبناء كبار و الجمعيات المباشرين، المشتركين ذلك في بما (1)'),
             (-14, 'المصدر: النشرة الإحصائية السنوية للمغرب2025'),
             (-15, 'Source : Annuaire Statistique du Maroc 2025'),
             (-17, 'عمالة : مكناس')],
  'after': [(2, 'ENERGIE'), (3, 'الطاقة'), (5, "11-2 Ventes d'électricité"), (6, '2-11 مبيعات الطاقة الكهربائية')]},
 {'original_index': 10751,
  'old_text': '770,9 millions de Kwh : Ventes de la préfecture de Meknès en électricité770,9 millions de Kwh : Ventes de la préfecture de Meknès en '
              'électricité9،770 مليون كيلوواط في الساعة: مبيعات عمالة مكناس من الطاقة الكهربائية9،770 مليون كيلوواط في الساعة: مبيعات عمالة مكناس من الطاقة '
              'الكهربائية',
  'new_text': '783,1 millions de Kwh : Ventes de la préfecture de Meknès en électricité783,1 millions de Kwh : Ventes de la préfecture de Meknès en '
              'électricité1،783 مليون كيلوواط في الساعة: مبيعات عمالة مكناس من الطاقة الكهربائية1،783 مليون كيلوواط في الساعة: مبيعات عمالة مكناس من الطاقة '
              'الكهربائية',
  'new_nodes': ['783,1',
                ' ',
                'millions',
                ' ',
                'de',
                ' ',
                'Kwh',
                ' ',
                ':',
                ' ',
                'Ventes de la préfecture de Meknès en ',
                'électricité',
                '783,1',
                ' ',
                'millions',
                ' ',
                'de',
                ' ',
                'Kwh',
                ' ',
                ':',
                ' ',
                'Ventes de la préfecture de Meknès en ',
                'électricité',
                '1',
                '،',
                '783',
                ' ',
                'مليون كيلوواط في',
                ' ',
                'الساعة',
                ':',
                ' ',
                'مبيعات',
                ' ',
                'عمالة',
                ' ',
                'مكناس',
                ' ',
                'من',
                ' ',
                'الطاقة',
                ' ',
                'الكهربائية',
                '1',
                '،',
                '783',
                ' ',
                'مليون كيلوواط في',
                ' ',
                'الساعة',
                ':',
                ' ',
                'مبيعات',
                ' ',
                'عمالة',
                ' ',
                'مكناس',
                ' ',
                'من',
                ' ',
                'الطاقة',
                ' ',
                'الكهربائية'],
  'before': [(-6, 'م : معطيات مصححة'),
             (-10, 'R :données rectifiées'),
             (-11, 'المصدر: النشرة الإحصائية السنوية للمغرب 2025'),
             (-15, 'Source : Annuaire Statistique du Maroc 2025')],
  'after': [(13, 'TOURISME'), (14, 'السياحة'), (18, 'Liste des Tableaux'), (19, 'لائحة الجداول')]},
 {'original_index': 10752,
  'old_text': '770,9 millions de Kwh : Ventes de la préfecture de Meknès en électricité',
  'new_text': '783,1 millions de Kwh : Ventes de la préfecture de Meknès en électricité',
  'new_nodes': ['783,1', ' ', 'millions', ' ', 'de', ' ', 'Kwh', ' ', ':', ' ', 'Ventes de la préfecture de Meknès en ', 'électricité'],
  'before': [(-7, 'م : معطيات مصححة'),
             (-11, 'R :données rectifiées'),
             (-12, 'المصدر: النشرة الإحصائية السنوية للمغرب 2025'),
             (-16, 'Source : Annuaire Statistique du Maroc 2025')],
  'after': [(12, 'TOURISME'), (13, 'السياحة'), (17, 'Liste des Tableaux'), (18, 'لائحة الجداول')]},
 {'original_index': 10753,
  'old_text': '770,9 millions de Kwh : Ventes de la préfecture de Meknès en électricité',
  'new_text': '783,1 millions de Kwh : Ventes de la préfecture de Meknès en électricité',
  'new_nodes': ['783,1', ' ', 'millions', ' ', 'de', ' ', 'Kwh', ' ', ':', ' ', 'Ventes de la préfecture de Meknès en ', 'électricité'],
  'before': [(-8, 'م : معطيات مصححة'),
             (-12, 'R :données rectifiées'),
             (-13, 'المصدر: النشرة الإحصائية السنوية للمغرب 2025'),
             (-17, 'Source : Annuaire Statistique du Maroc 2025')],
  'after': [(11, 'TOURISME'), (12, 'السياحة'), (16, 'Liste des Tableaux'), (17, 'لائحة الجداول')]},
 {'original_index': 10754,
  'old_text': '9،770 مليون كيلوواط في الساعة: مبيعات عمالة مكناس من الطاقة الكهربائية',
  'new_text': '1،783 مليون كيلوواط في الساعة: مبيعات عمالة مكناس من الطاقة الكهربائية',
  'new_nodes': ['1',
                '،',
                '783',
                ' ',
                'مليون كيلوواط في',
                ' ',
                'الساعة',
                ':',
                ' ',
                'مبيعات',
                ' ',
                'عمالة',
                ' ',
                'مكناس',
                ' ',
                'من',
                ' ',
                'الطاقة',
                ' ',
                'الكهربائية'],
  'before': [(-9, 'م : معطيات مصححة'),
             (-13, 'R :données rectifiées'),
             (-14, 'المصدر: النشرة الإحصائية السنوية للمغرب 2025'),
             (-18, 'Source : Annuaire Statistique du Maroc 2025')],
  'after': [(10, 'TOURISME'), (11, 'السياحة'), (15, 'Liste des Tableaux'), (16, 'لائحة الجداول')]},
 {'original_index': 10755,
  'old_text': '9،770 مليون كيلوواط في الساعة: مبيعات عمالة مكناس من الطاقة الكهربائية',
  'new_text': '1،783 مليون كيلوواط في الساعة: مبيعات عمالة مكناس من الطاقة الكهربائية',
  'new_nodes': ['1',
                '،',
                '783',
                ' ',
                'مليون كيلوواط في',
                ' ',
                'الساعة',
                ':',
                ' ',
                'مبيعات',
                ' ',
                'عمالة',
                ' ',
                'مكناس',
                ' ',
                'من',
                ' ',
                'الطاقة',
                ' ',
                'الكهربائية'],
  'before': [(-10, 'م : معطيات مصححة'),
             (-14, 'R :données rectifiées'),
             (-15, 'المصدر: النشرة الإحصائية السنوية للمغرب 2025'),
             (-19, 'Source : Annuaire Statistique du Maroc 2025')],
  'after': [(9, 'TOURISME'), (10, 'السياحة'), (14, 'Liste des Tableaux'), (15, 'لائحة الجداول')]},
 {'original_index': 10832,
  'old_text': '82 établissements touristiques classés 4 425 lits82 établissements touristiques classés 4 425 lits82 مؤسسة سياحية مصنفة 425 4 سرير82 مؤسسة '
              'سياحية مصنفة 425 4 سرير',
  'new_text': '81 établissements touristiques classés 4 347 lits81 établissements touristiques classés 4 347 lits81 مؤسسة سياحية مصنفة 347 4 سرير81 مؤسسة '
              'سياحية مصنفة 347 4 سرير',
  'new_nodes': ['81 établissements touristiques',
                ' ',
                'classés 4 347 lits',
                '81 établissements touristiques',
                ' ',
                'classés 4 347 lits',
                '81',
                ' ',
                'مؤسسة سياحية مصنفة ',
                '347',
                ' ',
                '4',
                ' ',
                'سرير',
                '81',
                ' ',
                'مؤسسة سياحية مصنفة ',
                '347',
                ' ',
                '4',
                ' ',
                'سرير'],
  'before': [(-8, 'المصدر: النشرة الإحصائية السنوية للمغرب2025'), (-9, 'Source : Annuaire Statistique du Maroc 2025'), (-11, 'عمالة : مكناس'), (-13, '2 088')],
  'after': [(10, 'TOURISME'), (11, 'السياحة'), (13, '12 -2 Les établissements classés'), (14, '12 - 2 المؤسسات المصنفة')]},
 {'original_index': 10834,
  'old_text': '82 établissements touristiques classés 4 425 lits',
  'new_text': '81 établissements touristiques classés 4 347 lits',
  'new_nodes': ['81 établissements touristiques', ' ', 'classés 4 347 lits'],
  'before': [(-10, 'المصدر: النشرة الإحصائية السنوية للمغرب2025'),
             (-11, 'Source : Annuaire Statistique du Maroc 2025'),
             (-13, 'عمالة : مكناس'),
             (-15, '2 088')],
  'after': [(8, 'TOURISME'), (9, 'السياحة'), (11, '12 -2 Les établissements classés'), (12, '12 - 2 المؤسسات المصنفة')]},
 {'original_index': 10836,
  'old_text': '82 établissements touristiques classés 4 425 lits',
  'new_text': '81 établissements touristiques classés 4 347 lits',
  'new_nodes': ['81 établissements touristiques', ' ', 'classés 4 347 lits'],
  'before': [(-12, 'المصدر: النشرة الإحصائية السنوية للمغرب2025'),
             (-13, 'Source : Annuaire Statistique du Maroc 2025'),
             (-15, 'عمالة : مكناس'),
             (-17, '2 088')],
  'after': [(6, 'TOURISME'), (7, 'السياحة'), (9, '12 -2 Les établissements classés'), (10, '12 - 2 المؤسسات المصنفة')]},
 {'original_index': 10838,
  'old_text': '82 مؤسسة سياحية مصنفة 425 4 سرير',
  'new_text': '81 مؤسسة سياحية مصنفة 347 4 سرير',
  'new_nodes': ['81', ' ', 'مؤسسة سياحية مصنفة ', '347', ' ', '4', ' ', 'سرير'],
  'before': [(-14, 'المصدر: النشرة الإحصائية السنوية للمغرب2025'),
             (-15, 'Source : Annuaire Statistique du Maroc 2025'),
             (-17, 'عمالة : مكناس'),
             (-19, '2 088')],
  'after': [(4, 'TOURISME'), (5, 'السياحة'), (7, '12 -2 Les établissements classés'), (8, '12 - 2 المؤسسات المصنفة')]},
 {'original_index': 10840,
  'old_text': '82 مؤسسة سياحية مصنفة 425 4 سرير',
  'new_text': '81 مؤسسة سياحية مصنفة 347 4 سرير',
  'new_nodes': ['81', ' ', 'مؤسسة سياحية مصنفة ', '347', ' ', '4', ' ', 'سرير'],
  'before': [(-16, 'المصدر: النشرة الإحصائية السنوية للمغرب2025'),
             (-17, 'Source : Annuaire Statistique du Maroc 2025'),
             (-19, 'عمالة : مكناس'),
             (-21, '2 088')],
  'after': [(2, 'TOURISME'), (3, 'السياحة'), (5, '12 -2 Les établissements classés'), (6, '12 - 2 المؤسسات المصنفة')]}]

def ptext(p):
    return "".join((t.text or "") for t in p.xpath(".//w:t", namespaces=NS))

def find_anchor(texts, text, start, stop):
    start = max(0, start)
    stop = min(len(texts), stop)
    for j in range(start, stop):
        if texts[j] == text:
            return j
    return None

def score_candidate(texts, candidate, spec):
    score = 0.0
    matched = 0

    for rel, text in spec["before"]:
        j = find_anchor(texts, text, candidate - 45, candidate)
        if j is not None:
            matched += 1
            actual = j - candidate
            score += 40.0 - min(20.0, abs(actual - rel))

    for rel, text in spec["after"]:
        j = find_anchor(texts, text, candidate + 1, candidate + 46)
        if j is not None:
            matched += 1
            actual = j - candidate
            score += 40.0 - min(20.0, abs(actual - rel))

    score -= abs(candidate - spec["original_index"]) * 0.0001
    return matched, score

def choose_candidate(texts, spec):
    old_text = spec["old_text"]
    new_text = spec["new_text"]

    candidates = []
    for i, txt in enumerate(texts):
        if txt == old_text:
            candidates.append((i, "old"))
        elif txt == new_text:
            candidates.append((i, "new"))

    if not candidates:
        raise RuntimeError(
            "Correction introuvable. "
            f"Ancien texte={old_text!r} / nouveau texte={new_text!r}"
        )

    if len(candidates) == 1:
        return candidates[0]

    ranked = []
    for i, state in candidates:
        matched, score = score_candidate(texts, i, spec)
        ranked.append((matched, score, -abs(i-spec["original_index"]), i, state))

    ranked.sort(reverse=True)
    best = ranked[0]

    if best[0] == 0:
        raise RuntimeError(
            f"Correction ambiguë pour {old_text!r} : "
            f"{len(candidates)} candidats et aucun contexte fiable."
        )

    return best[3], best[4]

def apply_spec(paragraphs, texts, spec):
    idx, state = choose_candidate(texts, spec)
    p = paragraphs[idx]

    if state == "new":
        return "already", idx

    nodes = p.xpath(".//w:t", namespaces=NS)
    new_nodes = spec["new_nodes"]

    if len(nodes) != len(new_nodes):
        raise RuntimeError(
            f"Structure de texte inattendue pour {spec['old_text']!r} "
            f"au paragraphe réel {idx} : {len(nodes)} segments au lieu de {len(new_nodes)}."
        )

    for node, new_text in zip(nodes, new_nodes):
        node.text = new_text

    texts[idx] = ptext(p)

    # Changing a parent paragraph containing text-box paragraphs can also
    # alter the text of descendant paragraphs. Refresh only a small local
    # neighborhood, not the whole document.
    for j in range(max(0, idx-6), min(len(paragraphs), idx+7)):
        texts[j] = ptext(paragraphs[j])

    if texts[idx] != spec["new_text"]:
        raise RuntimeError(
            f"Validation locale échouée au paragraphe réel {idx}."
        )

    return "changed", idx

def main():
    if not TARGET.exists():
        raise FileNotFoundError(f"Word final introuvable : {TARGET}")

    backup_dir = SORTIES / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = backup_dir / f"Annuaire_HCP_2025_FINAL_avant_synthese_{stamp}.docx"
    shutil.copy2(TARGET, backup)

    with zipfile.ZipFile(TARGET, "r") as zin:
        root = ET.fromstring(zin.read("word/document.xml"))
        paragraphs = root.xpath(".//w:p", namespaces=NS)

        if len(paragraphs) < 10000:
            raise RuntimeError(
                f"Structure Word trop courte : {len(paragraphs)} paragraphes détectés."
            )

        texts = [ptext(p) for p in paragraphs]

        changed = 0
        already = 0
        relocated = 0

        for spec in PATCHES:
            status, real_idx = apply_spec(paragraphs, texts, spec)

            if status == "changed":
                changed += 1
            else:
                already += 1

            if real_idx != spec["original_index"]:
                relocated += 1

        # Final semantic validation.
        for spec in PATCHES:
            idx, _state = choose_candidate(texts, spec)
            if texts[idx] != spec["new_text"]:
                raise RuntimeError(
                    f"Validation finale échouée pour {spec['new_text']!r}."
                )

        new_xml = ET.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        tmp = TARGET.with_name(TARGET.stem + "_TEMP.docx")

        with zipfile.ZipFile(tmp, "w") as zout:
            for item in zin.infolist():
                data = new_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)

    tmp.replace(TARGET)

    print("=" * 100)
    print("SYNCHRONISATION INDICATEURS / ENCADRES : OK")
    print("=" * 100)
    print(f"Corrections appliquées : {changed}")
    print(f"Déjà correctes : {already}")
    print(f"Corrections relocalisées dynamiquement : {relocated}")
    print(f"Word final : {TARGET}")
    print(f"Sauvegarde : {backup}")

if __name__ == "__main__":
    main()
