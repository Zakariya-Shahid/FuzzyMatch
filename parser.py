import pandas as pd
from fuzzywuzzy import fuzz
import openpyxl


# creating a function to get filter the url, title and id of the product
def filter():
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

    return dentalLeader, dentaltrey, donatalia, krugg, gerho, elident, umbra, vsdental


def masterSheet():
    krugg = filter()[3]
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


# function to write the dataframe to an excel file
def writeExcel(df):
    df.to_excel('output.xlsx', index=False)

# function to fuzzy match the title of all the dataframes with the title of krugg
def fuzzyMatch():
    # getting all the dataframes
    dentalLeaderdf = filter()[0]
    dentaltreydf = filter()[1]
    donataliadf = filter()[2]
    kruggdf = filter()[3]
    gerhodf = filter()[4]
    elidentdf = filter()[5]
    umbradf = filter()[6]
    vsdentaldf = filter()[7]
    maindf = masterSheet()
    print("Master sheet created")
    # adding the url column of all the dataframes to the main dataframe based on the fuzzy match
    for i in range(len(maindf)):
        print(i)
        for j in range(len(dentalLeaderdf)):
            if fuzz.ratio(maindf['Title'][i], dentalLeaderdf['title'][j]) > 90:
                maindf['dentalLeader'][i] = dentalLeaderdf['url'][j]
        for k in range(len(dentaltreydf)):
            if fuzz.ratio(maindf['Title'][i], dentaltreydf['Title'][k]) > 90:
                maindf['dentaltrey'][i] = dentaltreydf['URL'][k]
        for l in range(len(donataliadf)):
            if fuzz.ratio(maindf['Title'][i], donataliadf['Title'][l]) > 90:
                maindf['donatalia'][i] = donataliadf['URL'][l]
        for m in range(len(gerhodf)):
            if fuzz.ratio(maindf['Title'][i], gerhodf['Title'][m]) > 90:
                maindf['gerho'][i] = gerhodf['URL'][m]
        for n in range(len(elidentdf)):
            if fuzz.ratio(maindf['Title'][i], elidentdf['Title'][n]) > 90:
                maindf['elident'][i] = elidentdf['URL'][n]
        for o in range(len(umbradf)):
            if fuzz.ratio(maindf['Title'][i], umbradf['Title'][o]) > 90:
                maindf['umbra'][i] = umbradf['URL'][o]
        for p in range(len(vsdentaldf)):
            if fuzz.ratio(maindf['Title'][i], vsdentaldf['title'][p]) > 90:
                maindf['vsdental'][i] = vsdentaldf['URL'][p]

        writeExcel(maindf)


print("Starting...")
fuzzyMatch()
