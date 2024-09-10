import streamlit as st
import xlrd
import os
from io import StringIO
import pandas as pd
import progressbar

PASSWORD = "Swati@IMR8180"



def part1(keyword):
    ch1 = ("<strong>\nChapter 1: Introduction</strong>\n\t1.1 Scope and Coverage"
          +"\n\n<strong>Chapter 2:Executive Summary</strong>")
    return ch1.replace("\n","<br />").replace("\t","&emsp;")

def part1B(keyword):
    ch1 = ("\n\n<strong>Chapter 3: Market Landscape</strong>\n\t3.1 Market Dynamics\n\t\t3.1.1 Drivers\n\t\t3.1.2 Restraints\n\t\t3.1.3 Opportunities"
           + "\n\t\t3.1.4 Challenges\n\t3.2 Market Trend Analysis\n\t3.3 PESTLE Analysis\n\t3.4 Porter's Five Forces Analysis"
           + "\n\t3.5 Industry Value Chain Analysis \n\t3.6 Ecosystem \n\t3.7 Regulatory Landscape"
           + "\n\t3.8 Price Trend Analysis \n\t3.9 Patent Analysis\n\t3.10 Technology Evolution"
           + "\n\t3.11 Investment Pockets\n\t3.12 Import-Export Analysis")
    return ch1.replace("\n", "<br />").replace("\t", "&emsp;")

def part2(keyword,segmentation):
    count= 4
    segment = segmentation.split("#")
    ch2 = ""
    segment_heading = []
    print(keyword)
    
    for s in segment:
        x = s.split("::")
        segment_heading.append(x[0])
        ch2 += ("\n\n<strong>Chapter "+str(count)+": "+str(keyword)+" Market by "+x[0]+"</strong>\n\t"+str(count)+".1 "+str(keyword)+" Market Snapshot and Growth Engine\n\t"+str(count)+".2 "+str(keyword)+" Market Overview")
        y = x[1].split(",")
        counter = 3
        for types in y:
            ch2 += ("\n\t"+str(count)+"."+str(counter)+" "+str(types)+"\n\t\t"+str(count)+"."+str(counter)+".1 Introduction and Market Overview\n\t\t"
                    +str(count)+"."+str(counter)+".2 Historic and Forecasted Market Size in Value USD and Volume Units (2017-2032F)\n\t\t"
                    +str(count)+"."+str(counter)+".3 Key Market Trends, Growth Factors and Opportunities\n\t\t"
                   +str(count)+"."+str(counter)+".4 "+str(types)+": Geographic Segmentation Analysis")
            counter = counter + 1
        count = count + 1
    return ch2.replace("\n","<br />").replace("\t","&emsp;"),segment_heading,count

def part3(keyword, players, count):
    t = players.split(",")
    ch3 = ("\n\n<strong>Chapter " + str(count) + ": Company Profiles and Competitive Analysis</strong>\n\t" 
           + str(count) + ".1 Competitive Landscape\n\t\t"
           + str(count) + ".1.1 Competitive Benchmarking\n\t\t"
           + str(count) + ".1.2 " + str(keyword) + " Market Share by Manufacturer (2023)\n\t\t"
           + str(count) + ".1.3 Industry BCG Matrix\n\t\t"
           + str(count) + ".1.4 Heat Map Analysis\n\t\t"
           + str(count) + ".1.5 Mergers and Acquisitions\n\t\t")

    counter = 2
    for i in t:
        if counter == 2:
            ch3 += ("\n\t" + str(count) + "." + str(counter) + " " + i.strip().upper()
                    + "\n\t\t" + str(count) + "." + str(counter) + ".1 Company Overview"
                    + "\n\t\t" + str(count) + "." + str(counter) + ".2 Key Executives"
                    + "\n\t\t" + str(count) + "." + str(counter) + ".3 Company Snapshot"
                    + "\n\t\t" + str(count) + "." + str(counter) + ".4 Role of the Company in the Market"
                    + "\n\t\t" + str(count) + "." + str(counter) + ".5 Sustainability and Social Responsibility"
                    + "\n\t\t" + str(count) + "." + str(counter) + ".6 Operating Business Segments"
                    + "\n\t\t" + str(count) + "." + str(counter) + ".7 Product Portfolio"
                    + "\n\t\t" + str(count) + "." + str(counter) + ".8 Business Performance"
                    + "\n\t\t" + str(count) + "." + str(counter) + ".9 Key Strategic Moves and Recent Developments"
                    + "\n\t\t" + str(count) + "." + str(counter) + ".10 SWOT Analysis")
        else:
            ch3 += ("\n\t" + str(count) + "." + str(counter) + " " + i.strip().upper())
        counter += 1
    
    return ch3.replace("\n", "<br />").replace("\t", "&emsp;"), count

def part4(keyword,regions,count,segmentation):
    segment = segmentation.split("#")
    t = regions.split(",")
    ch4 = "\n\n<strong>Chapter "+str(count)+":"+str(keyword)+" Market Analysis, Insights and Forecast, 2016-2028</strong>\n\t"+str(count)+".1 Market Overview"
    counter = 2
    #for s in segment:
      #  x = s.split("::")
       # ch4 = ch4 + "\n\t"+str(count)+"."+str(counter)+" Historic and Forecasted Market Size By "+x[0]
        #y = x[1].split(",")
        #for idx,z in enumerate(y):
           # ch4 = ch4 + "\n\t\t"+str(count)+"."+str(counter)+"."+str(idx+1)+" "+str(z)
        #counter = counter + 1
    
   # return ch4.replace("\n","<br />").replace("\t","&emsp;"),count
    for i in t:
        counter = 1
        region = i.split("(")
        #ch4 = ch4 + "\n\n<strong>Chapter "+str(count)+": "+region[0].strip()+" "+str(keyword)+" Market Analysis, Insights and Forecast, 2016-2028</strong>"
        ch4 = ch4 + ("\n\t"+str(count)+"."+str(counter)+" Key Market Trends, Growth Factors and Opportunities"
                     +"\n\t"+str(count)+"."+str(counter+1)+" Impact of Covid-19"
                     +"\n\t"+str(count)+"."+str(counter+2)+" Key Players"
                     +"\n\t"+str(count)+"."+str(counter+3)+" Key Market Trends, Growth Factors and Opportunities")
        segment_counter = 4
        for s in segment:
            x = s.split("::")
            ch4 = ch4 + "\n\t"+str(count)+"."+str(segment_counter)+" Historic and Forecasted Market Size By "+x[0]
            y = x[1].split(",")
            for idx,z in enumerate(y):
                ch4 = ch4 + "\n\t\t"+str(count)+"."+str(segment_counter)+"."+str(idx+1)+" "+str(z)
            segment_counter = segment_counter + 1
            
        region[1] = region[1].replace(")","")
        ch4 = ch4 +"\n\t"+str(count)+"."+str(segment_counter)+" Historic and Forecast Market Size by Country"
        for idx, country in enumerate(region[1].split(",")[0].split(";")):    
            ch4 = ch4 +"\n\t\t"+str(count)+"."+str(segment_counter)+"."+str(idx+1)+" "+country.strip();
        counter+=1
        count = count + 1
    
    return ch4.replace("\n","<br />").replace("\t","&emsp;"),count


def part6(keyword, count):
    count = count 
    ch4 = (
        "\n\n<strong>Chapter " + str(count) + " Analyst Viewpoint and Conclusion</strong>\n"
        + str(count) + ".1 Recommendations and Concluding Analysis\n"
        + str(count) + ".2 Potential Market Strategies\n"
        "\n<strong>Chapter " + str(count + 1) + " Research Methodology</strong>\n"
        + str(count + 1) + ".1 Research Process\n"
        + str(count + 1) + ".2 Primary Research\n"
        + str(count + 1) + ".3 Secondary Research\n\n"
    )
    return ch4.replace("\n", "<br />").replace("\t", "&emsp;")

def table(keyword,segmentation,regions,companies):
    
    segment = segmentation.split("#")
    ch2 = "<strong>LIST OF TABLES</strong>\n\n"
    segment_heading = []
    
    ch2 += ("TABLE 001. EXECUTIVE SUMMARY\nTABLE 002. XYZ MARKET BARGAINING POWER OF SUPPLIERS\nTABLE 003. XYZ MARKET BARGAINING POWER OF CUSTOMERS\n"
            +"TABLE 004. XYZ MARKET COMPETITIVE RIVALRY\nTABLE 005. XYZ MARKET THREAT OF NEW ENTRANTS\nTABLE 006. XYZ MARKET THREAT OF SUBSTITUTES")  
    count = 7
    for s in segment:
        x = s.split("::")
        segment_heading.append(x[0])
        ch2 += ("\nTABLE "+str(count).zfill(3)+". XYZ MARKET BY "+x[0].upper().strip())
        y = x[1].split(",")
        count = count + 1
        for types in y:
            ch2 += ("\nTABLE "+str(count).zfill(3)+". "+str(types).upper().strip()+" MARKET OVERVIEW (2016-2030)")
            count = count + 1
            
    t = regions.split(",")
    
    for i in t:
        for y in segment_heading:
            ch2 += ("\nTABLE "+str(count).zfill(3)+". "+str(i).upper().strip()+" XYZ MARKET, BY "+str(y).upper().strip()+" (2016-2030)")
            count = count + 1
        ch2 += ("\nTABLE "+str(count).zfill(3)+". "+str(i[0]).upper().strip()+" XYZ MARKET, BY COUNTRY (2016-2030)")
        count = count + 1
    
    c = companies.split(",")
    
    for z in c:
        if "Others" not in z or "and Others" not in z:
            ch2+= ("\nTABLE "+str(count).zfill(3)+". "+z.upper().strip()+": SNAPSHOT")
            count = count+1
            ch2+= ("\nTABLE "+str(count).zfill(3)+". "+z.upper().strip()+": BUSINESS PERFORMANCE")
            count = count+1
            ch2+= ("\nTABLE "+str(count).zfill(3)+". "+z.upper().strip()+": PRODUCT PORTFOLIO")
            count = count+1
            ch2+= ("\nTABLE "+str(count).zfill(3)+". "+z.upper().strip()+": KEY STRATEGIC MOVES AND DEVELOPMENTS")
            
    ch2 = ch2.replace("XYZ",keyword.strip().upper())
    
    return ch2.replace("\n","<br />").replace("\t","&emsp;")


def create_html_report(uploaded_file, output_dir):
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Error reading the Excel file: {e}")
            return

        required_columns = ['Keyword', 'Segmentation', 'Regions', 'Players']
        for col in required_columns:
            if col not in df.columns:
                st.error(f"Missing required column: {col}")
                return
        
        keyword = df['Keyword'].iloc[0]
        segmentation = df['Segmentation'].iloc[0]
        regions = df['Regions'].iloc[0]
        players = df['Players'].iloc[0]

        ch1_content = part1(keyword)
        ch1b_content = part1B(keyword)
        part2_content, segment_heading, count = part2(keyword, segmentation)
        ch3_content, count = part3(keyword, players, count)
        part4_content, count = part4(keyword, regions, count, segmentation)
        part6_content, count=  part6(keyword, count)      
        table_content = table(keyword, segmentation, regions, players)

        html_content = ch1_content + ch1b_content + part2_content + ch3_content + part4_content + part6_content + table_content


        try:
            os.makedirs(output_dir, exist_ok=True)
            output_file = os.path.join(output_dir, 'report.html')
            with open(output_file, 'w') as f:
                f.write(html_content)
            st.write(f"Report saved to {output_file}")
        except Exception as e:
            st.error(f"Error saving the report: {e}")

def figures(keyword,segmentation,regions,companies):
    
    segment = segmentation.split("#")
    ch2 = "<strong>LIST OF FIGURES</strong>\n\n"
    segment_heading = []
    
    ch2 += ("FIGURE 001. YEARS CONSIDERED FOR ANALYSIS\nFIGURE 002. SCOPE OF THE STUDY\nFIGURE 003. XYZ MARKET OVERVIEW BY REGIONS\n"
            +"FIGURE 004. PORTER'S FIVE FORCES ANALYSIS\nFIGURE 005. BARGAINING POWER OF SUPPLIERS\nFIGURE 006. COMPETITIVE RIVALRY"
            +"FIGURE 007. THREAT OF NEW ENTRANTS\nFIGURE 008. THREAT OF SUBSTITUTES\nFIGURE 009. VALUE CHAIN ANALYSIS\nFIGURE 010. PESTLE ANALYSIS")  
    count = 11
    for s in segment:
        x = s.split("::")
        segment_heading.append(x[0])
        ch2 += ("\nFIGURE "+str(count).zfill(3)+". XYZ MARKET OVERVIEW BY "+x[0].upper().strip())
        y = x[1].split(",")
        count = count + 1
        for types in y:
            ch2 += ("\nFIGURE "+str(count).zfill(3)+". "+str(types).upper().strip()+" MARKET OVERVIEW (2016-2030)")
            count = count + 1
            
    t = regions.split(",")
    
    for i in t:
        ch2 += ("\nFIGURE "+str(count).zfill(3)+". "+str(i).upper().strip()+" XYZ MARKET OVERVIEW BY COUNTRY (2016-2030)")
        count = count + 1
    
            
    ch2 = ch2.replace("XYZ",keyword.strip().upper())
    
    return ch2.replace("\n","<br />").replace("\t","&emsp;")

def process_excel_file(uploaded_file, id_keyword, id_segmentation, id_player, id_region, regions):
    wb = xlrd.open_workbook(file_contents=uploaded_file.read())
    sheet = wb.sheet_by_index(0)
    
    ans_toc = ""
    ans_fig = ""
    bar = progressbar.ProgressBar(maxval=sheet.nrows, widgets=[progressbar.Bar('=', '[', ']'), ' ', progressbar.Percentage()])
    bar.start()
    
    for j in range(1, sheet.nrows):
        keyword = sheet.cell_value(j, id_keyword).strip()
        segmentation = sheet.cell_value(j, id_segmentation)
        players = sheet.cell_value(j, id_player)
        
        p1 = part1(keyword)
        p1b = part1B(keyword)
        p2, segment_heading, count = part2(keyword, segmentation)
        p3, count = part3(keyword, players, count)
        p4 = part4(keyword, regions, count, segmentation)
        p6 = part6(keyword, count)
        ans_fig += table(keyword, segmentation, regions, players) + "<br /><br />"
        ans_fig += figures(keyword, segmentation, regions, players) + "\n"
        ans_toc += p1 + p1b + p2[0] + p3[0] + p4 + p6
        
        bar.update(j + 1)
    
    bar.finish()
    return ans_toc, ans_fig


if __name__ == "__main__":
    pass

def main():
    st.title("Region TOC Generator")

    st.sidebar.header("Upload Excel File")
    uploaded_file = st.sidebar.file_uploader("Choose a file", type="xls", key="file_uploader_1")
    
    if uploaded_file is not None:
        # Load the Excel file
        wb = xlrd.open_workbook(file_contents=uploaded_file.read())
        sheet = wb.sheet_by_index(0)

        st.sidebar.header("Column IDs")
        id_segmentation = st.sidebar.number_input("Segmentation Column ID", min_value=0, value=12)
        id_player = st.sidebar.number_input("Player Column ID", min_value=0, value=8)
        id_keyword = st.sidebar.number_input("Keyword Column ID", min_value=0, value=1)
        id_region = st.sidebar.number_input("Region Column ID", min_value=0, value=5)
        
        regions = st.text_input("Regions", "North America, Eastern Europe, Western Europe, Asia Pacific, Middle East & Africa, South America")
        
        progress_bar = st.progress(0)
        ans_toc = ""
        ans_fig = ""

        for j in range(1, sheet.nrows):
            keyword = sheet.cell_value(j, id_keyword).strip()
            segmentation = sheet.cell_value(j, id_segmentation)
            players = sheet.cell_value(j, id_player)

        p1 = part1(sheet.cell_value(j,id_keyword).strip())
        p1b = part1B(sheet.cell_value(j,id_keyword).strip())
        p2 = part2(sheet.cell_value(j,id_keyword).strip(),sheet.cell_value(j,id_segmentation))
#         p3a = part3A(sheet.cell_value(j,id_keyword).strip(),p2[1])
        p3 = part3(sheet.cell_value(j,id_keyword).strip(),sheet.cell_value(j,id_player),p2[2])
        #p4 = part4(sheet.cell_value(j,id_keyword).strip(),sheet.cell_value(j,id_region),p3[1]+1,sheet.cell_value(j,id_segmentation))
        #p5 = part5(sheet.cell_value(j,id_keyword).strip(),sheet.cell_value(j,id_region),p4[1]+1,sheet.cell_value(j,id_segmentation))
        # p4 = part4(sheet.cell_value(j,id_keyword).strip(),p3[1])
        p4 = part4(sheet.cell_value(j,id_keyword).strip(),sheet.cell_value(j,id_region),p3[1]+1,sheet.cell_value(j,id_segmentation))
        p6 = part6(sheet.cell_value(j,id_keyword).strip(),p4[1])
        
        ans_fig = ans_fig + '"' + table(sheet.cell_value(j,id_keyword).strip(),sheet.cell_value(j,id_segmentation),regions,sheet.cell_value(j,id_player)) + "<br /><br />"
        
        ans_fig = ans_fig +  figures(sheet.cell_value(j,id_keyword).strip(),sheet.cell_value(j,id_segmentation),regions,sheet.cell_value(j,id_player)) + '"\n'
        
        ans_toc = ans_toc + '"' + str(p1) + str(p1b) + str(p2[0]) + str(p3[0]) + str(p4) + str(p6) +'"\n'
            
        progress_bar.progress((j + 1) / sheet.nrows)

        st.header("Table of Contents")
        st.write(ans_toc)

        st.header("Figures")
        st.write(ans_fig)

        # Create download buttons for the text files
        toc_text = StringIO(ans_toc)
        fig_text = StringIO(ans_fig)
        
        st.sidebar.download_button(
            label="Download Table of Contents",
            data=toc_text.getvalue(),
            file_name='toc.txt',
            mime='text/plain'
        )
        
        st.sidebar.download_button(
            label="Download Figures",
            data=fig_text.getvalue(),
            file_name='figures.txt',
            mime='text/plain'
        )

if __name__ == "__main__":
    password = st.sidebar.text_input("Password", type="password")

    # Check password
    if password == PASSWORD:
        main()  # If password is correct, run the app
    else:
        st.error("Incorrect password. Please try again.")
