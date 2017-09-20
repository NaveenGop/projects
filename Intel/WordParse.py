import os, olefile, re, tempfile
from joblib import Parallel, delayed
from datetime import date


def create_tsv(di, file):
    file.write('%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n' % (di['name'].decode(), di['prod code'].decode(), di['sku'].decode(), di['density'].decode(), di['sku_letter'].decode(), di['form_factor'].decode(), di['pba'].decode(), di['save time'], di['path']))


def process_bigdata(di: dict) -> list:
    s = (re.sub(b'\\r|\\n', b'', di['big data'])).split(b')')
    del di['big data']
    s = [x for x in s if b'XXX' in x]
    lst = []
    for item in s:
        temp = {'name': di['name'], 'prod code': di['prod code'], 'path': di['path'], 'save time': di['save time']}
        if b'=' in item:
            b = item.partition(b'=')
            j = b[2].lstrip()
            temp['pba'] = j + b')'
            temp['sku_letter'] = (b[0][-3:]).strip(b'- ')
            j = (j[j.find(b'(')+1:])
            a = re.search(b'[\\d.]{2,4}[\\s]*[GT][\\s]*[Bb]*[\\s]*', j)
            try:
                k = re.sub(b'[GgBbTt\\s]', b'', j[a.start():a.end()])
                if b'.' in k:
                    temp['density'] = (k.replace(b'.', b'P'))[0:]
                else:
                    temp['density'] = k
                temp['form_factor'] = j.replace(j[a.start():a.end()], b'')
            except AttributeError:
                temp['density'] = b''
                temp['form_factor'] = j
            temp['sku'] = temp['density'] + temp['sku_letter']
        elif b'Alternative' in item:
            temp['pba'] = item.lstrip() + b')'
            temp['sku_letter'] = b'Alternate'
            temp['density'] = b'Alternate'
            temp['sku'] = b'Alternate'
            temp['form_factor'] = b'Alternate'
        lst.append(temp)
    return lst


def proper_reg(temp_text_path: str):
    """Finds the name and prod code. Big data contains the other essential information. Big data is processed by process."""
    with open(temp_text_path, 'r+b') as f:
        text = f.read()
    try:
        dit = {}
        boolean = True
        if "2d label spec" in temp_text_path.lower():
            i = text.find(b'PP= PRODUCT IDENTIFIER') + 22
        else:
            i = text.find(b'PP = Product Identifier')+23
            boolean = False
        regex = re.compile(b'\\s*=\\s*')
        j = regex.search(text, i)
        if abs(i-j.start()) <= 3:
            dit['prod code'] = text[i:i+3].strip()
            regex = re.compile(b'\\r|\\n')
            dit['name'] = text[j.end(): regex.search(text, j.end()).start()].strip().title()
        else:
            dit['prod code'] = text[j.end():j.end()+2].strip()
            dit['name'] = text[i: j.start()].strip().title()
        if boolean:
            i = text.find(b'PBA number', j.end()) + 11
            dit['big data'] = text[i:text.find(b'Human readable', i)-2]
            if b'human' in dit['name']:
                dit['name'] = dit['name'][:dit['name'].find(b'human')]
        else:
            i = text.find(b'PBA#)', j.end()) + 5
            dit['big data'] = text[i:text.find(b'CC', i) - 2]
        return dit
    except:
        print(temp_text_path.split('/')[-1], ' is formatted irregularly. Was not processed')


def format_doc(enter: str, tempdir: str):
    """Opens a .doc path, creates a temp text file and directory, and writes only the alphanumeric and punctuation characters to the file."""
    ole = olefile.OleFileIO(enter)
    txt = ole.openstream('WordDocument').read()
    temp_txt_path = tempdir + "\\" + enter.split('/')[-1] + ".txt"
    with open(temp_txt_path, 'w+b') as f:
        f.write(re.sub(b'[^\\x20-\\x7F\\r]', b'', txt[txt.find(b'REV'): txt.rfind(b'Barcode')]))
    return temp_txt_path


def extract_metadata(path: str) -> dict:
    """Finds the path and last saved time for each doc."""
    di = {}
    ole = olefile.OleFileIO(path)
    meta = ole.get_metadata()
    di['path'] = path
    di['save time'] = date(meta.last_saved_time.year, meta.last_saved_time.month, meta.last_saved_time.day)
    return di


def rev(prod: list, revis: str):
    try:
        i = 0
        r = len(prod)
        while i < r:
            r = len(prod)
            g = i + 1
            tup = prod[i].partition(revis)
            rev1 = int(tup[2].lstrip()[0:2])
            a = True
            while g < r:
                r = len(prod)
                uptup = prod[g].partition(revis)
                uprev = int(uptup[2].lstrip()[0:2])
                if tup[0].rstrip().rsplit(' ', 1)[1] == uptup[0].rstrip().rsplit(' ', 1)[1]:
                    if rev1 > uprev:
                        prod.remove(prod[g])
                    elif uprev > rev1:
                        prod.remove(prod[i])
                        a = False
                        break
                    else:
                        g += 1
                else:
                    g += 1
            if a:
                i += 1
    except IndexError:
        print(prod, i, g, r)


def find_duplicates(parallel, tempdir, *arg) -> list:
    """First for loop creates a dictionary for each PBA for all the doc paths and appends it to final_set. Second half deletes duplicate dictionaries."""
    final_set = []
    for prod_doc in arg:
        if len(prod_doc) > 0:
            rev(prod_doc, "rev")
            meta = parallel(delayed(extract_metadata)(enter) for enter in prod_doc)
            prod_doc = parallel(delayed(format_doc)(enter, tempdir) for enter in prod_doc)
            if len(prod_doc) >= 1:
                prod_doc = parallel(delayed(proper_reg)(enter) for enter in prod_doc)
                for i in range(len(meta)):
                    prod_doc[i] = {**prod_doc[i], **meta[i]}
                prod_doc = parallel(delayed(process_bigdata)(di) for di in prod_doc)
            for lst in prod_doc:
                for di in lst:
                    final_set.append(di)
    r = len(final_set)
    f = 0
    while f < r:
        r = len(final_set)
        g = f+1
        di = final_set[f]
        temp = {'prod code': di['prod code'], 'sku': di['sku'], 'density': di['density'], 'pba': di['pba'].partition(b'(')[0]}
        while g < r:
            temp2 = {'prod code': final_set[g]['prod code'], 'sku': final_set[g]['sku'], 'density': final_set[g]['density'], 'pba': final_set[g]['pba'].partition(b'(')[0]}
            if temp == temp2:
                final_set.remove(final_set[g])
                r = len(final_set)
            else:
                g += 1
        f += 1
    return final_set


def sort_list(doc_name: str, *vart):
    """Checks if a string contains the identifiers in vart. If it does return the initial string, if not return None."""
    for var in vart:
        if var.lower() in doc_name.lower() and "(old" not in doc_name and "~" not in doc_name:
            return doc_name


def find_docs(path: str, vart: tuple = ()) -> list:
    """Scans a directory for .doc files with the identifiers in vart. Calls itself to scan subdirectories."""
    doc_list = []
    for sub_entry in os.scandir(path):
        if sub_entry.is_dir():
            if len(vart) != 0:
                for var in vart:
                    if var.lower() in sub_entry.name.lower():
                        doc_list += (find_docs(path + "/" + sub_entry.name, vart))
        elif sub_entry.name.endswith(".doc"):
            doc_list.append(path + "/" + sub_entry.name)
    return doc_list


def scan_dir(path: str, *vart) -> list:
    """"Scans a directory for .doc files with the identifiers in vart. Calls find_docs to scan subdirectories."""
    doc_list = []
    for entry in os.scandir(path):
        if entry.is_dir():
            if len(vart) != 0:
                for var in vart:
                    if var.lower() in entry.name.lower():
                        doc_list += (find_docs(path + "/" + entry.name, vart))
            else:
                doc_list += (find_docs(path + "/" + entry.name))
        elif '.doc' in entry.name:
            doc_list.append(path + "/" + entry.name)
    return list(set(doc_list))


def main(file_name='lenis.tsv', path="//amr/ec/proj/fm/nsg/NAND/Shares/FailureAnalysis/rfrickey/20160808_Label_Spec_File_Copy"):
    """ """
    folder_list = os.scandir(path)
    f = open(file_name, 'w+')
    with tempfile.TemporaryDirectory() as tempdir:
        with Parallel(n_jobs=-1) as parallel:
            for orig_fold in folder_list:
                prod_doc = scan_dir(path + "/" + orig_fold.name, 'label spec', orig_fold.name)
                prod_doc = parallel(delayed(sort_list)(entry, "isn label spec", "2d label spec") for entry in prod_doc)
                prod_doc = [x for x in prod_doc if x is not None]
                final_set = find_duplicates(parallel, tempdir, prod_doc)
                for i in final_set:
                    create_tsv(i, f)
    f.close()

if __name__ == '__main__':
    main()
