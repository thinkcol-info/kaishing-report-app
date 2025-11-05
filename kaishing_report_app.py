import os
from datetime import datetime, timedelta
import io

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dotenv import load_dotenv
from io import BytesIO
from tempfile import NamedTemporaryFile
from docx import Document
from docx.shared import Inches
from send_report_enhanced import create_word_document, create_interactive_html
import boto3


st.set_page_config(page_title="KaiShing OAK Report", layout="wide")


def get_data_from_dynamodb(table_name, aws_key, aws_secret, aws_region):
    try:
        dynamodb = boto3.resource(
            'dynamodb',
            aws_access_key_id=aws_key,
            aws_secret_access_key=aws_secret,
            region_name=aws_region,
        )
        table = dynamodb.Table(table_name)
        response = table.scan()
        items = response['Items']
        while 'LastEvaluatedKey' in response:
            response = table.scan(ExclusiveStartKey=response['LastEvaluatedKey'])
            items.extend(response['Items'])
        return pd.DataFrame(items)
    except Exception as e:
        st.error(f"Error fetching {table_name}: {e}")
        return pd.DataFrame()


def filter_df_by_range(df, start_date, end_date, date_col='createdAt'):
    if df.empty or date_col not in df.columns:
        return df
    numeric = pd.to_numeric(df[date_col], errors='coerce')
    dt = pd.to_datetime(numeric, unit='s', errors='coerce').dt.tz_localize('UTC')
    df = df.copy()
    df[date_col] = dt
    mask = (df[date_col] >= start_date) & (df[date_col] <= end_date)
    return df[mask].copy()


def build_figures_and_render(account_df, usage_df, askai_df):
    total_accounts = len(account_df)
    subscription_counts = account_df['subscription_level'].value_counts(dropna=False)
    pro_users = subscription_counts.get('pro', 0)
    team_users = subscription_counts.get('team', 0)

    st.subheader("Overall Platform Statistics")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Active Accounts", f"{total_accounts}")
    c2.metric("Pro Tier Users", f"{pro_users}")
    c3.metric("Team Tier Users", f"{team_users}")

    # Figures to reuse for exports
    fig_wau = go.Figure()
    fig_heatmap = go.Figure()

    if not usage_df.empty:
        hkt = 'Asia/Hong_Kong'
        usage_df = usage_df.copy()
        usage_df['createdAt_HKT'] = usage_df['createdAt'].dt.tz_convert(hkt)
        wau_df = usage_df.groupby(pd.Grouper(key='createdAt_HKT', freq='W-Mon'))['account'].nunique().reset_index()
        wau_df.rename(columns={'account': 'Weekly Active Users'}, inplace=True)
        fig_wau = px.line(
            wau_df,
            x='createdAt_HKT',
            y='Weekly Active Users',
            title='<b>Weekly Active Users (WAU) Trend</b>',
            labels={'createdAt_HKT': 'Week (in HKT)'},
            markers=True,
            template='plotly_white',
        )
        fig_wau.update_layout(title_x=0.5)
        fig_wau.update_traces(line_color='#1f77b4')
        st.plotly_chart(fig_wau, use_container_width=True)

        usage_df['day_of_week'] = usage_df['createdAt_HKT'].dt.day_name()
        usage_df['hour_of_day'] = usage_df['createdAt_HKT'].dt.hour
        activity_heatmap_data = usage_df.pivot_table(
            index='hour_of_day', columns='day_of_week', values='id', aggfunc='count'
        ).fillna(0)
        day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        activity_heatmap_data = activity_heatmap_data.reindex(columns=day_order, fill_value=0)
        zmax = activity_heatmap_data.values.max() if activity_heatmap_data.values.size > 0 else 0
        fig_heatmap = go.Figure(
            data=go.Heatmap(
                z=activity_heatmap_data.values,
                x=activity_heatmap_data.columns,
                y=activity_heatmap_data.index,
                colorscale=[
                    [0.0, '#eef5ff'],
                    [0.2, '#d6e9ff'],
                    [0.4, '#9ecae1'],
                    [0.6, '#6baed6'],
                    [0.8, '#3182bd'],
                    [1.0, '#08519c']
                ],
                zauto=False,
                zmin=0,
                zmax=zmax,
            )
        )
        fig_heatmap.update_layout(
            title='<b>User Activity Heatmap (by Day and Hour)</b>',
            xaxis_title='Day of the Week',
            yaxis_title='Hour of the Day (HKT)',
            title_x=0.5,
        )
        st.plotly_chart(fig_heatmap, use_container_width=True)

    fig_site_activity = go.Figure()
    if not usage_df.empty and 'site_code' in usage_df.columns:
        activity_by_site = usage_df['site_code'].value_counts().reset_index()
        activity_by_site.columns = ['site_code', 'action_count']
        fig_site_activity = px.treemap(
            activity_by_site,
            path=[px.Constant("All Sites"), 'site_code'],
            values='action_count',
            title='<b>Activity Distribution by Site Code</b>',
            template='plotly_white',
            hover_data={'action_count': ':.0f'},
            color_continuous_scale='Greens',
        )
        fig_site_activity.update_layout(title_x=0.5, margin=dict(t=50, l=25, r=25, b=25))
        fig_site_activity.update_traces(textinfo="label+value", textfont_size=14)
        st.plotly_chart(fig_site_activity, use_container_width=True)

    fig_features = go.Figure()
    if not usage_df.empty and 'usage_type' in usage_df.columns:
        feature_usage_counts = usage_df['usage_type'].value_counts().reset_index()
        feature_usage_counts.columns = ['feature', 'count']
        fig_features = px.bar(
            feature_usage_counts,
            x='count',
            y='feature',
            orientation='h',
            title='<b>What Features Are Users Exploring?</b>',
            template='plotly_white',
            text='count',
        )
        fig_features.update_yaxes(categoryorder='total ascending')
        fig_features.update_layout(title_x=0.5, xaxis_title='Number of Times Feature Was Used')
        fig_features.update_traces(textposition='outside', marker_color='#5a4fcf')
        st.plotly_chart(fig_features, use_container_width=True)

    fig_askai_sites = go.Figure()
    if askai_df is not None and not askai_df.empty and 'site_code' in askai_df.columns:
        ask_ai_by_site = askai_df['site_code'].value_counts().reset_index()
        ask_ai_by_site.columns = ['site_code', 'query_count']
        fig_askai_sites = px.bar(
            ask_ai_by_site,
            x='query_count',
            y='site_code',
            orientation='h',
            title='<b>AskAI Pioneers: Adoption by Site</b>',
            text='query_count',
            template='plotly_white',
            color='query_count',
            color_continuous_scale=px.colors.sequential.Viridis,
        )
        fig_askai_sites.update_yaxes(categoryorder='total ascending')
        fig_askai_sites.update_layout(title_x=0.5, xaxis_title='Number of AskAI Queries', yaxis_title='Site Code', coloraxis_showscale=False)
        st.plotly_chart(fig_askai_sites, use_container_width=True)

    figures = {
        'fig_wau': fig_wau,
        'fig_heatmap': fig_heatmap,
        'fig_site_activity': fig_site_activity,
        'fig_features': fig_features,
        'fig_askai_sites': fig_askai_sites,
        'fig_askai_keywords': go.Figure(),
    }

    return (
        figures,
        {
            'total_accounts': total_accounts,
            'pro_users': pro_users,
            'team_users': team_users,
        },
    )


def main():
    load_dotenv(override=True)

    st.title("KaiShing OAK Report Generator")
    st.caption("Generate interactive reports and export downloads. End users can choose time periods.")

    with st.sidebar:
        st.header("Configuration")
        start = st.date_input("Start date", value=(datetime.now() - timedelta(days=30)).date())
        end = st.date_input("End date", value=datetime.now().date())

        if end < start:
            st.error("End date must be after start date.")
            st.stop()

        sections = st.multiselect(
            "Sections",
            options=["overview", "engagement", "adoption", "features", "askai"],
            default=["overview", "engagement", "adoption", "features", "askai"],
        )

        st.markdown("---")
        st.subheader("Exports")
        st.caption("Use the buttons in the main page to download Word, HTML, and Summary Excel.")

    aws_key = os.getenv("KAISHING_DYNAMODB_ACCESS_KEY_ID")
    aws_secret = os.getenv("KAISHING_DYNAMODB_SECRET_ACCESS_KEY")
    aws_region = os.getenv("KAISHING_DYNAMODB_REGION")

    with st.spinner("Loading data from DynamoDB..."):
        account_df = get_data_from_dynamodb("oak-account-ks", aws_key, aws_secret, aws_region)
        usage_df = get_data_from_dynamodb("oak-usage-log-ks", aws_key, aws_secret, aws_region)
        askai_df = get_data_from_dynamodb("oak-ask-ai-ks", aws_key, aws_secret, aws_region)
        transcription_df = get_data_from_dynamodb("oak-transcription-ks", aws_key, aws_secret, aws_region)

    users_to_exclude = ['kian.so@thinkcol.com', 'hetty.pun@thinkcol.com', 'adawan@kaishing.com.hk']
    site_code_map = {
        'eddiecheuk@kaishing.com.hk': 'HQ-IT', 'ksitsupport@kaishing.com.hk': 'HQ-IT', 'aegeancoast@kaishing.com.hk': 'AC',
        'dacychung@kaishing.com.hk': 'ICC', 'lewislam@kaishing.com.hk': 'ICC', 'Vcity@kaishing.com.hk': 'VCY',
        'yohomidtown@kaishing.com.hk': 'YMT', 'leightonhill@supreme-mgt.com.hk': 'LH', 'riva@supreme-mgt.com.hk': 'RV',
        'tpmm@kaishing.com.hk': 'TPMM', 'palmsprings@kaishing.com.hk': 'PS', 'castello@kaishing.com.hk': 'CAS',
        'newtown3@kaishing.com.hk': 'NTP3R', 'millencity@kaishing.com.hk': 'M388', 'mounthaven@kaishing.com.hk': 'MH',
        'victorwong@supreme-mgt.com.hk': 'UMA', 'epc@kaishing.com.hk': 'EPC-C', 'millencity5@kaishing.com.hk': 'MMC418',
        'apm@kaishing.com.hk': 'MMC418', 'taipocentre@kaishing.com.hk': 'TPC', 'parkisland@kaishing.com.hk': 'PI',
        'thewings3a@kaishing.com.hk': 'TW3A', 'pacificview@kaishing.com.hk': 'PV', 'cffy@chifufayuen.hk': 'CFFY',
        '98hms@kaishing.com.hk': '98HMS', 'stanford@kaishing.com.hk': 'SFV', 'lepalais@kaishing.com.hk': 'LPS',
        'avignon@kaishing.com.hk': 'AGN', 'pmt@kaishing.com.hk': 'PMT', 'mountregency@kaishing.com.hk': 'MR',
        'somerset@kaishing.com.hk': 'SOM', 'emilyho@kaishing.com.hk': 'NTPI', 'garychan@kaishing.com.hk': 'CAC',
        'dynastycourt@kaishing.com.hk': 'DC', 'eastpoint@kaishing.com.hk': 'EPCR', 'grandyoho@kaishing.com.hk': 'GYR',
        'hillsborough@kaishing.com.hk': 'HC', 'kodakhouse11@kaishing.com.hk': 'KHII', 'oceanwings@kaishing.com.hk': 'OW',
        'pokfulam@kaishing.com.hk': 'PG', 'royalpalms@kaishing.com.hk': 'RP', 'concerto@kaishing.com.hk': 'VC',
        'brownieyu@kaishing.com.hk': 'AFFC', 'hlypm@kaishing.com.hk': 'HLY', 'thewings2@kaishing.com.hk': 'TW2',
        'mayfair@kaishing.com.hk': 'MG', 'affc@kaishing.com.hk': 'AFFC', 'villabythepark@kaishing.com.hk': 'VP',
        'celestecourt@kaishing.com.hk': 'CC', 'ls@kaishing.com.hk': 'LS', 'suntuenmun@kaishing.com.hk': 'STMC',
        'lagrove@kaishing.com.hk': 'LG', 'yohowest@wespire.com.hk': 'YOW', 'yohohouse@wespire.com.hk': 'YOW',
        'kennedy38@supreme-mgt.com.hk': 'K38', 'homantinhill@supreme-mgt.com.hk': 'HMT', 'landmarkn@kaishing.com.hk': 'LN',
        'metroplaza@kaishing.com.hk': 'MP', 'yohomall-1@kaishing.com.hk': 'YM1', 'ylplaza@kaishing.com.hk': 'YLP',
        'kingspark@kaishing.com.hk': 'KPV', 'candicewong@kaishing.com.hk': 'KCC', 'rhapsody@kaishing.com.hk': 'VR',
        'lgar@kaishing.com.hk': 'LGAR', 'rseacrest@kaishing.com.hk': 'RSC', 'yukpocourt@kaishing.com.hk': 'YPC',
        'villaathena@kaishing.com.hk': 'VA', 'vincenttse@supreme-mgt.com.hk': 'VY'
    }

    # Clean and enrich
    account_df = account_df[~account_df['account'].isin(users_to_exclude)].copy() if not account_df.empty else account_df
    if not usage_df.empty:
        usage_df = usage_df[~usage_df['account'].isin(users_to_exclude)].copy()
        usage_df['site_code'] = usage_df['account'].map(site_code_map).fillna('Unknown')
    if not askai_df.empty:
        askai_df = askai_df[~askai_df['account'].isin(users_to_exclude)].copy()
        askai_df['site_code'] = askai_df['account'].map(site_code_map).fillna('Unknown')

    # Filter by selected date range
    start_dt = pd.to_datetime(start)
    end_dt = pd.to_datetime(datetime.combine(end, datetime.max.time()))
    usage_filtered = filter_df_by_range(usage_df, start_dt.tz_localize('UTC'), end_dt.tz_localize('UTC')) if not usage_df.empty else usage_df
    askai_filtered = filter_df_by_range(askai_df, start_dt.tz_localize('UTC'), end_dt.tz_localize('UTC')) if not askai_df.empty else askai_df
    transcription_filtered = filter_df_by_range(transcription_df, start_dt.tz_localize('UTC'), end_dt.tz_localize('UTC')) if not transcription_df.empty else transcription_df

    # Render charts
    figures, kpi_vals = build_figures_and_render(account_df, usage_filtered, askai_filtered)

    st.markdown("---")
    st.subheader("Downloads")

    # Summary Excel (account-level metrics) built from usage, transcription, and askai
    def build_summary_df(acc_df, usage_df, transcription_df, ask_df):
        if acc_df is None or acc_df.empty:
            return pd.DataFrame()
        accounts = acc_df[['account']].drop_duplicates().rename(columns={'account': 'account'})

        def counter_by_usage_type(df, usage_type_value, name):
            """Count actions by usage_type from usage_df"""
            if df is None or df.empty or 'usage_type' not in df.columns:
                return pd.DataFrame(columns=['account', name])
            filtered = df[df['usage_type'] == usage_type_value] if usage_type_value else df
            c = filtered.groupby('account').size().reset_index(name=name)
            return c

        def counter_from_transcription(df, name):
            """Count transcriptions from transcription table"""
            if df is None or df.empty:
                return pd.DataFrame(columns=['account', name])
            if 'account' not in df.columns:
                return pd.DataFrame(columns=['account', name])
            c = df.groupby('account').size().reset_index(name=name)
            return c

        def counter_from_askai(df, name):
            """Count AskAI questions"""
            if df is None or df.empty:
                return pd.DataFrame(columns=['account', name])
            c = df.groupby('account').size().reset_index(name=name)
            return c

        # Count generated transcripts (from transcription table - initial generation)
        gen_trans = counter_from_transcription(transcription_df, 'generated transcripts')

        # Count regenerated transcripts (usage_type contains 'regenerate' and 'transcript')
        # Check for common usage_type patterns
        if not usage_df.empty and 'usage_type' in usage_df.columns:
            regen_transcript_patterns = usage_df[usage_df['usage_type'].str.contains('regenerate.*transcript', case=False, na=False)]
            if not regen_transcript_patterns.empty:
                regen_trans = regen_transcript_patterns.groupby('account').size().reset_index(name='regenerated transcripts')
            else:
                regen_trans = pd.DataFrame(columns=['account', 'regenerated transcripts'])
        else:
            regen_trans = pd.DataFrame(columns=['account', 'regenerated transcripts'])

        # Count initial summaries (usage_type contains 'initial' and 'summary')
        if not usage_df.empty and 'usage_type' in usage_df.columns:
            init_sum_patterns = usage_df[usage_df['usage_type'].str.contains('initial.*summary|generate.*summary', case=False, na=False) &
                                         ~usage_df['usage_type'].str.contains('regenerate', case=False, na=False)]
            if not init_sum_patterns.empty:
                init_sum = init_sum_patterns.groupby('account').size().reset_index(name='initial summaries')
            else:
                init_sum = pd.DataFrame(columns=['account', 'initial summaries'])
        else:
            init_sum = pd.DataFrame(columns=['account', 'initial summaries'])

        # Count regenerated summaries (usage_type contains 'regenerate' and 'summary')
        if not usage_df.empty and 'usage_type' in usage_df.columns:
            regen_sum_patterns = usage_df[usage_df['usage_type'].str.contains('regenerate.*summary', case=False, na=False)]
            if not regen_sum_patterns.empty:
                regen_sum = regen_sum_patterns.groupby('account').size().reset_index(name='regenerated summaries')
            else:
                regen_sum = pd.DataFrame(columns=['account', 'regenerated summaries'])
        else:
            regen_sum = pd.DataFrame(columns=['account', 'regenerated summaries'])

        # Count generated notes (usage_type contains 'generate' or 'initial' and 'note', but not 'regenerate')
        if not usage_df.empty and 'usage_type' in usage_df.columns:
            gen_notes_patterns = usage_df[usage_df['usage_type'].str.contains('generate.*note|initial.*note', case=False, na=False) &
                                          ~usage_df['usage_type'].str.contains('regenerate', case=False, na=False)]
            if not gen_notes_patterns.empty:
                gen_notes = gen_notes_patterns.groupby('account').size().reset_index(name='generated notes')
            else:
                gen_notes = pd.DataFrame(columns=['account', 'generated notes'])
        else:
            gen_notes = pd.DataFrame(columns=['account', 'generated notes'])

        # Count regenerated notes (usage_type contains 'regenerate' and 'note')
        if not usage_df.empty and 'usage_type' in usage_df.columns:
            regen_notes_patterns = usage_df[usage_df['usage_type'].str.contains('regenerate.*note', case=False, na=False)]
            if not regen_notes_patterns.empty:
                regen_notes = regen_notes_patterns.groupby('account').size().reset_index(name='regenerated notes')
            else:
                regen_notes = pd.DataFrame(columns=['account', 'regenerated notes'])
        else:
            regen_notes = pd.DataFrame(columns=['account', 'regenerated notes'])

        # Count AskAI questions
        ask_cnt = counter_from_askai(ask_df, 'askai questions')

        out = accounts.merge(gen_trans, on='account', how='left') \
                      .merge(regen_trans, on='account', how='left') \
                      .merge(init_sum, on='account', how='left') \
                      .merge(regen_sum, on='account', how='left') \
                      .merge(gen_notes, on='account', how='left') \
                      .merge(regen_notes, on='account', how='left') \
                      .merge(ask_cnt, on='account', how='left')
        for col in out.columns:
            if col != 'account':
                out[col] = out[col].fillna(0).astype(int)
        # Try to add username if present
        if 'username' in acc_df.columns:
            out = out.merge(acc_df[['account','username']], on='account', how='left')
            cols = ['account','username'] + [c for c in out.columns if c not in ['account','username']]
            out = out[cols]
        return out

    summary_df = build_summary_df(account_df, usage_filtered, transcription_filtered, askai_filtered)
    excel_buffer = None
    if not summary_df.empty:
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            summary_df.to_excel(writer, index=False, sheet_name='Summary')
        excel_buffer.seek(0)
        st.download_button(
            label="Download Summary Excel",
            data=excel_buffer,
            file_name=f"kaishing_summary_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # Export HTML using the same builder for email visuals
    if st.button("Build HTML Report"):
        with NamedTemporaryFile(suffix='.html', delete=False) as tmp_html:
            tmp_html_path = tmp_html.name
        # create_interactive_html writes to file path
        create_interactive_html(
            usage_log_df=usage_filtered,
            ask_ai_df=askai_filtered,
            selected_sections=sections,
            total_accounts=kpi_vals['total_accounts'],
            pro_users=kpi_vals['pro_users'],
            team_users=kpi_vals['team_users'],
            output_path=tmp_html_path,
        )
        with open(tmp_html_path, 'rb') as f:
            html_bytes = f.read()
        st.download_button(
            label="Download HTML Report",
            data=html_bytes,
            file_name=f"kaishing_report_{datetime.now().strftime('%Y%m%d')}.html",
            mime="text/html",
        )

    # Export Word using the same visuals (generate temporary file then provide download)
    if st.button("Build Word Report"):
        with NamedTemporaryFile(suffix='.docx', delete=False) as tmp_docx:
            tmp_docx_path = tmp_docx.name
        create_word_document(
            figures_dict=figures,
            selected_sections=sections,
            total_accounts=kpi_vals['total_accounts'],
            pro_users=kpi_vals['pro_users'],
            team_users=kpi_vals['team_users'],
            usage_log_valid_time_df=usage_filtered,
            ask_ai_df=askai_filtered,
            output_path=tmp_docx_path,
        )
        with open(tmp_docx_path, 'rb') as f:
            word_bytes = f.read()
        st.download_button(
            label="Download Word Report",
            data=word_bytes,
            file_name=f"kaishing_report_{datetime.now().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    st.caption("Tip: Use the sidebar to adjust dates and sections, then download fresh exports.")


if __name__ == "__main__":
    main()
