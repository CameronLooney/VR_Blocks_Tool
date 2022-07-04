import pandas as pd
import datetime
import streamlit as st
from st_aggrid import AgGrid

import plotly.express as px

st.set_page_config(page_title="VR Reporting Tool", layout="wide")
def main():
    st.title('VR Block Report Builder')
    upload_report = st.sidebar.file_uploader("Upload VR Report", type="xlsx")
    if st.sidebar.button('Generate Report'):
        if upload_report:
            data = pd.read_excel(upload_report, sheet_name=0, engine="openpyxl")
            def data_processing_vr(df):
                df = df[["cust_name", "sales_order","lob","mpn","supply_comments","supply_ETA","OrderCommitHigh","delivery_block_item"]]
                df = df[df['delivery_block_item'] == "VR"]
                df['OrderCommitHigh'] = pd.to_datetime(df['OrderCommitHigh'], dayfirst=True)
                return df


            df_date_passed = data_processing_vr(data)
            def find_high_commit_rows(df):
                df = df[datetime.datetime.now() > df['OrderCommitHigh']]
                df = df[df['supply_comments'].isnull()]
                return df

            df_date_passed= find_high_commit_rows( df_date_passed )
            st.subheader("Order Commit Date has passed without Comment")
            AgGrid(df_date_passed)

            def upcoming_orders(df):
                ten_days = datetime.datetime.now() + datetime.timedelta(10)
                df = df.loc[(df['OrderCommitHigh'] < ten_days) & (df['OrderCommitHigh'] > datetime.datetime.now())]
                df = df[df['supply_comments'].isnull()]
                return df
            x =data_processing_vr(data)
            x = upcoming_orders(x)
            st.subheader("Order Commit Date is within the next week without comment")
            AgGrid(x, editable=True,copyHeadersToClipboard=True)



            def vr_blocks_by_lob():
                df = data_processing_vr(data)
                fig = px.histogram(data_frame=df, x="lob",color = "lob", title='VR Blocks by LOB', text_auto=True,

                                   )
                fig.update_xaxes(categoryorder="category ascending")
                st.plotly_chart(fig)

            vr_blocks_by_lob()

            def vr_blocks_date_passed():
                df = data_processing_vr(data)
                df["Order Commit High Date"] = df.apply(lambda row: "Date Passed" if datetime.datetime.now() > row['OrderCommitHigh'] else "Date Upcoming",axis=1)
                fig = px.histogram(data_frame=df, x="lob", color="Order Commit High Date", title='VR Blocks - Order Commit Date Passed',
                                   text_auto=True,

                                   )
                fig.update_xaxes(categoryorder="category ascending")
                st.plotly_chart(fig)
            vr_blocks_date_passed()

            def vr_blocks_commentary():
                df = pd.read_excel(upload_report, sheet_name=0, engine="openpyxl")
                df = df[df['delivery_block_item'] == "VR"]
                print(df["supply_comments"])
                # if supply_comments is not null then set df["holder"] to "yes" and if not set to "no"
                df["Commentary"] = df["supply_comments"].isnull()
                # if df["holder"] is true then set df["holder"] to "yes" and if not set to "no"
                df["Commentary"] = df["Commentary"].map({True: "No Comment", False: "Comment"})

                fig = px.histogram(data_frame=df, x="lob", color="Commentary", title='VR Blocks - Commentary',
                                   text_auto=True

                                   )
                fig.update_xaxes(categoryorder="category ascending")
                st.plotly_chart(fig)
            vr_blocks_commentary()






            def excel(df1,df2):
                import io
                buffer = io.BytesIO()
                writer = pd.ExcelWriter(buffer, date_format='yyyy-mm-dd', datetime_format='yyyy-mm-dd')
                df1.to_excel(writer, sheet_name='Date Passed Orders', index=False)
                df2.to_excel(writer, sheet_name='Upcoming Date Orders', index=False)
                worksheet1 = writer.sheets['Date Passed Orders']
                (max_row, max_col) = df1.shape
                worksheet1.set_column(1, max_col, 20)
                worksheet2 = writer.sheets['Upcoming Date Orders']
                (max_row, max_col) = df2.shape
                worksheet2.set_column(1, max_col, 20)
                writer.save()
                return buffer

            to_excel = excel(df_date_passed, x)






            # download file
            def download(buffer):
                st.download_button(
                    label="Download VR Report",
                    data=buffer,
                    file_name="VR_Report.xlsx",
                    mime="application/vnd.ms-excel"
                )

            download(to_excel)



if __name__ == "__main__":
    main()