#basic script to read the output of a downloaded bibtex file from NASA ADS
#and strip out all one's co-authors, and dump to a spreadsheet for
#copy/pasting into an NSF co-authors spreadsheet.  It tries to convert
#all special characters to standard text, and indexes based on last
#name and first initial (so you should check your output if it fails).
#To get the input file (export-bibtex.bib) do a search on ADS, click on the
#"export" button, and export to bibtex.  Download the file, and run away.
#DEFINITELY check the output.  There are corner cases that aren't quite
#handled correctly (e.g. ADS sometimes seems to put line breaks inside a name)
#so caveat emptor.

import xlsxwriter
import argparse


def find_arg(line,start='{',stop='}'):
    istart=line.find(start)
    if istart<0:
        return None
    ncur=1
    icur=istart+1
    for i in range(icur,len(line)):
        if line[i]==start:
            ncur=ncur+1
        if line[i]==stop:
            ncur=ncur-1
        if ncur==0:
            return line[icur:i],i

def find_next_key(key,line,start='{',stop='}',offset=0):
    ind=line.find(key,offset)
    istart=line.find(start,ind)
    if istart<0:
        return None
    #print(ind,istart)
    icur=istart+1
    ncur=1
    for i in range(icur,len(line)):
        if line[i]==start:
            ncur=ncur+1
        if line[i]==stop:
            ncur=ncur-1
        if ncur==0:
            return line[istart+1:i],i
    print('unterminated key.  returning end of line.')
    return line[istart+1:],-1

def split_line(line,tag=' and '):
    keys=[]
    ii=0
    ll=line
    while ll.find(tag)>=0:
        ind=ll.find(tag)
        if ind>0:
            key=ll[:ind]
            keys.append(key)
            ll=ll[ind+len(tag):]
    return keys


def strip_special(str,start='{',stop='}'):
    while True:
        i1=str.find(start)
        if i1==-1:
            return str
        else:
            inside,ind=find_arg(str)
            head=str[:i1]
            foot=str[ind+1:]
            while inside.find(start)>=0:  #there are sometimes nested special characters
                inside,ind=find_arg(inside)
            inside=inside[-1]
            str=head+inside+foot
    return str
            
def strip_special_old(str,start='{',stop='}'):
    while True:
        i1=str.find(start)
        if i1==-1:
            return str
        else:
            i2=str.find(stop)
            if (i2<=i1):
                print('bad key: ',str)
            assert(i2>i1)
            if i1==0:
                head=''
            else:
                head=str[:i1]
            if i2==len(str)-1:
                tail=''
            else:
                tail=str[i2+1:]
            i3=i2-1
            while str[i3]==stop:
                i3=i3-1
            mid=str[i3]
            str=head+mid+tail

def parse_author(name):
    name=name.replace('~',' ')
    surname,ind=find_arg(name)
    prename=name[ind+2:]
    surname=surname.strip()
    prename=prename.strip()
    #print(surname)
    #print(prename)
    surname=strip_special(surname)
    prename=strip_special(prename)
    

    #print(surname,prename)

    #i1=name.find('{')
    #i2=name.rfind('}')
    #surname=name[i1+1:i2]
    #prename=name[i2+2:]
    #surname=surname.strip()
    #prename=prename.strip()
    #surname=strip_special(surname)
    #prename=strip_special(prename)
    return surname,prename
    

if True:
    f=open('export-bibtex.bib')
    lines=f.read()
    f.close()
    lines=lines.replace('\n',' ')
else:
    f=open('export-bibtex.bib')
    lines=f.readlines()
    f.close()
    for i in range(len(lines)):
        lines[i]=lines[i].strip()
    tot=''
    for line in lines:
        tot=tot+line
    lines=tot

if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument("-m","--min",nargs='+',type=int,default=0,help="Minimum number of authors per paper.",dest='nmin')
    parser.add_argument("-M","--max",nargs='+',type=int,default=99999,help="Maximum number of authors per paper.",dest='nmax')
    parser.add_argument("-f","--file",nargs='+',type=str,default='authors.xlsx',help='Output file name',dest='file')
    args=parser.parse_args()
    nmin=args.nmin
    nmax=args.nmax
    ofile=args.file
    if isinstance(ofile,list):
        ofile=ofile[0]
    if isinstance(nmin,list):
        nmin=nmin[0]
    if isinstance(nmax,list):
        nmax=nmax[0]
    ii=0
    authors={}
    npaper=0
    nused=0
    while True:
        ans=find_next_key('author',lines,offset=ii)
        if ans is None:
            print('stopping after ',npaper,' papers of which ',nused,' were parsed.')
            break
        else:
            tag=ans[0]
            ii=ans[1]
            npaper=npaper+1
        keys=split_line(tag)
        
        nauthor=len(keys)
        if (nauthor>=nmin)&(nauthor<nmax):
            nused=nused+1
            for key in keys:
                surname,prename=parse_author(key)
                if len(prename)>0:
                    nm=surname+' '+prename[0]
                if nm in authors.keys():
                    if len(prename)>len(authors[nm]):
                        authors[nm]=prename
                else:
                    authors[nm]=prename
                    


    workbook = xlsxwriter.Workbook(ofile)
    worksheet = workbook.add_worksheet()
    
    keys=authors.keys()

    for i,key in enumerate(keys):
        init=key[-1]
        surname=key[:-1].strip()
        name_use=surname+', '+authors[key]
        worksheet.write('A'+repr(i+1),name_use)
    #worksheet.write('B'+repr(i+1),init)
    #worksheet.write('C'+repr(i+1),authors[key])
    
    workbook.close()
