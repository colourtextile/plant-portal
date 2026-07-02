# --- DASHBOARD LOGIC (Updated for 0-Hide & Party Summary) ---
        else:
            st.markdown("<h2 style='color: #1F4E79; font-weight: bold; margin-top: -10px;'>📊 Live Analytics Dashboard</h2>", unsafe_allow_html=True)
            
            # --- 📅 1. CHOOSE FILTER RANGE ---
            st.markdown("### 🎛️ Filter Range Selection")
            filter_type = st.radio("📅 Select Dashboard Filter Range:", ["☀️ Day-Wise Filter", "📆 Month-Wise Filter"], horizontal=True)
            
            filtered_df_by_range = pd.DataFrame()
            display_range_label = ""
            
            if excel_loaded and not df.empty:
                df['Item Type'] = df['Item Type'].astype(str).str.upper().str.strip()
                df['Party Name'] = df['Party Name'].astype(str).str.strip()
                df['Date'] = df['Date'].astype(str).str.strip()
                df['parsed_date'] = pd.to_datetime(df['Date'], format='%d-%m-%Y', errors='coerce')
                df['Month_Year'] = df['parsed_date'].dt.strftime('%B %Y')
            
            if filter_type == "☀️ Day-Wise Filter":
                selected_filter_date = st.date_input("Choose Specific Date to View Summary", datetime.now())
                formatted_filter_date = selected_filter_date.strftime('%d-%m-%Y')
                display_range_label = f"SELECTED DATE: {formatted_filter_date}"
                if excel_loaded and not df.empty:
                    filtered_df_by_range = df[df['Date'] == formatted_filter_date]
            else:
                current_month_year = datetime.now().strftime('%B %Y')
                available_months = sorted(list(df['Month_Year'].dropna().unique())) if excel_loaded else []
                if current_month_year not in available_months: available_months.append(current_month_year)
                selected_month = st.selectbox("Choose Month and Year to View Summary", available_months)
                display_range_label = f"SELECTED MONTH: {selected_month.upper()}"
                if excel_loaded and not df.empty:
                    filtered_df_by_range = df[df['Month_Year'] == selected_month]

            # Aggregating values
            if not filtered_df_by_range.empty:
                item_groups = filtered_df_by_range.groupby('Item Type')['Total Pcs'].sum()
                party_groups = filtered_df_by_range.groupby('Party Name')['Total Pcs'].sum()
                
                # Zero-Value Filter: Sirf wo jinki value > 0 hai
                item_groups = item_groups[item_groups > 0]
                party_groups = party_groups[party_groups > 0]
            else:
                item_groups = pd.Series()
                party_groups = pd.Series()

            col_left, col_right = st.columns([1, 1])
            
            # --- ITEM PIECES SUMMARY ---
            with col_left:
                st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
                st.markdown(f"### 🏭 ITEM PIECES SUMMARY\n**{display_range_label}**")
                if not item_groups.empty:
                    st.table(item_groups.reset_index().rename(columns={'Item Type': 'Item', 'Total Pcs': 'Pcs'}))
                    st.success(f"TOTAL PRODUCED: {int(item_groups.sum()):,} Pcs")
                else:
                    st.info("No production data available for this range.")
                st.markdown('</div>', unsafe_allow_html=True)

            # --- PARTY WISE SUMMARY (Naya Section) ---
            with col_right:
                st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
                st.markdown(f"### 📦 PARTY WISE SUMMARY\n**{display_range_label}**")
                if not party_groups.empty:
                    st.table(party_groups.reset_index().rename(columns={'Party Name': 'Party', 'Total Pcs': 'Pcs'}))
                    st.info(f"TOTAL DISPATCH: {int(party_groups.sum()):,} Pcs")
                else:
                    st.info("No party dispatch data available.")
                st.markdown('</div>', unsafe_allow_html=True)
