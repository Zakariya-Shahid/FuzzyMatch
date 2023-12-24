import pandas as pd
from fuzzywuzzy import fuzz
from concurrent.futures import ThreadPoolExecutor



# getting all the columns of all dataframes
dentalLeader = pd.read_excel('dental_leader.xlsx')
dentaltrey = pd.read_excel('dentaltrey.xlsx')
donatalia = pd.read_excel('dontalia.it.xlsx')
gerho = pd.read_excel('gerho.it_v2.xlsx')
krugg = pd.read_excel('krugg.xlsx')
elident = pd.read_excel('elidentgroup.itv2.xlsx')
umbra = pd.read_excel('umbra.it_v2.xlsx')
vsdental = pd.read_excel('vsdental.xlsx')
dentalLeader = dentalLeader[['url', 'title', 'id']]
dentaltrey = dentaltrey[['URL', 'Title', 'ID']]
donatalia = donatalia[['URL', 'Title', 'Variant Manufacture ID']]
krugg = krugg[['URL', 'Title', 'SKU']]
gerho = gerho[['URL', 'Title', 'ID']]
elident = elident[['URL', 'Title', ' Variant ID']]
umbra = umbra[['URL', 'Title', 'ID']]
vsdental = vsdental[['URL', 'title', 'product_id']]


def masterSheet():
    # creating a new dataframe to store ID, title and url of krugg and urls of all other dataframes
    df = pd.DataFrame()
    df['ID'] = krugg['SKU']
    df['Title'] = krugg['Title']
    df['URL'] = krugg['URL']
    df['dentalLeader'] = ''
    df['dentaltrey'] = ''
    df['donatalia'] = ''
    df['gerho'] = ''
    df['elident'] = ''
    df['umbra'] = ''
    df['vsdental'] = ''

    return df

maindf = masterSheet()

# function to fuzzy match the title of all the dataframes with the title of krugg
def fuzzyMatch(start, end):
    i = start
    global maindf, dentalLeader, dentaltrey, donatalia, krugg, gerho, elident, umbra, vsdental
    print("Master sheet created")
    # adding the url column of all the dataframes to the main dataframe based on the fuzzy match
    while i<end:
        print(i,',')
        for j in range(len(dentalLeader)):
            if fuzz.ratio(maindf['Title'][i], dentalLeader['title'][j]) > 90:
                maindf['dentalLeader'][i] = dentalLeader['url'][j]
                print('added')
        for k in range(len(dentaltrey)):
            if fuzz.ratio(maindf['Title'][i], dentaltrey['Title'][k]) > 90:
                maindf['dentaltrey'][i] = dentaltrey['URL'][k]
        for l in range(len(donatalia)):
            if fuzz.ratio(maindf['Title'][i], donatalia['Title'][l]) > 90:
                maindf['donatalia'][i] = donatalia['URL'][l]
        for m in range(len(gerho)):
            if fuzz.ratio(maindf['Title'][i], gerho['Title'][m]) > 90:
                maindf['gerho'][i] = gerho['URL'][m]
        for n in range(len(elident)):
            if fuzz.ratio(maindf['Title'][i], elident['Title'][n]) > 90:
                maindf['elident'][i] = elident['URL'][n]
        for o in range(len(umbra)):
            if fuzz.ratio(maindf['Title'][i], umbra['Title'][o]) > 90:
                maindf['umbra'][i] = umbra['URL'][o]
        for p in range(len(vsdental)):
            if fuzz.ratio(maindf['Title'][i], vsdental['title'][p]) > 90:
                maindf['vsdental'][i] = vsdental['URL'][p]
        i+=1


print("Starting...")
executer = ThreadPoolExecutor(max_workers=32)
i=0
while i <= 20000:
    future = executer.submit(fuzzyMatch, i, i+10)
    i += 10

executer.shutdown(wait=True)
maindf.to_csv('masterSheet.csv', index=False)
print("Done")

