import streamlit as st
import pandas as pd
import json
from io import BytesIO
import os
import uuid

# ---------- helper: Streamlit ≥1.27 / pre‑1.27 -------------------------------
def safe_rerun(): (st.rerun if hasattr(st, "rerun") else st.experimental_rerun)()

# ---------- constants --------------------------------------------------------
OPTIONS_FILE = "qual_options.json"
FIRST_RUN_FLAG = "first_run_flag.txt"
DEFAULT_OPTIONS = {
    "Dominance": ["Exclusive", "Primary", "Secondary", "Passing reference"],
    "Prominence": ["Headline", "Co-visual Image"],
    "Spokesperson": ["Authored", "Interview", "Quote", "Mention"],
    "Page": [],
    "Tonality": ["Positive", "Negative", "Neutral"],
    "Category": ["Innovation", "Market share", "Leadership", "Customer relation",
                 "M&A", "Business Growth", "Products & Services", "Vision", "Work Environment"],
    "SavedUserCategories": []
}
MANDATORY = ["Dominance", "Prominence", "Spokesperson", "Page", "Tonality", "Category"]

# ---------- option bank helpers ---------------------------------------------
def initialize_bank():
    is_first_run = not os.path.exists(FIRST_RUN_FLAG)
    if is_first_run:
        json.dump(DEFAULT_OPTIONS, open(OPTIONS_FILE, "w"), indent=2)
        with open(FIRST_RUN_FLAG, "w") as f:
            f.write("First run completed")
        return DEFAULT_OPTIONS.copy()
    else:
        try:
            loaded = json.load(open(OPTIONS_FILE))
            for key in DEFAULT_OPTIONS:
                if key not in loaded:
                    loaded[key] = DEFAULT_OPTIONS[key]
                if key == "Category":
                    loaded[key] = DEFAULT_OPTIONS[key]
                else:
                    if not isinstance(loaded[key], list):
                        loaded[key] = DEFAULT_OPTIONS[key]
            return loaded
        except Exception as e:
            st.error(f"Error loading {OPTIONS_FILE}: {e}. Using default options. Please fix the JSON file.")
            return DEFAULT_OPTIONS.copy()

def save_bank(b):
    json.dump(b, open(OPTIONS_FILE, "w"), indent=2)

# Initialize the options bank
bank = initialize_bank()

# ---------- session-state bootstrap -----------------------------------------
init_vals = {
    "df_raw": pd.DataFrame(),
    "df_work": pd.DataFrame(),
    "qualified": pd.DataFrame(),
    "partial": pd.DataFrame(),
    "deleted": pd.DataFrame(),
    "to_be_decided": pd.DataFrame(),
    "row_ptr": 0,
    "bucket_row_ptr": 0,
    "total": 0,
    "file_uploaded": False,
    "current_category_index": 0,
    "selected_categories": [],
    "category_qualifications": [],
    "qualification_started": False,
    "saved_user_categories": bank.get("SavedUserCategories", []),
    "category_selection_order": [],
    "qualifications_by_category": {},
    "qualified_categories_by_row": {},
    "show_caution_message": False,
    "confirm_categories": False,
    "preview_bucket": None,
    "no_more_records_message": None,
}
for k, v in init_vals.items():
    st.session_state.setdefault(k, v)

# ---------- title / upload ---------------------------------------------------
st.title("📰 News Qualification App")

up = st.file_uploader(
    "Upload an **Excel** file (.xlsx / .xls)",
    type=["xlsx", "xls"],
    accept_multiple_files=False
)

# Process as soon as a file is chosen
if up and not st.session_state.file_uploaded:
    try:
        st.session_state.df_raw = pd.read_excel(up)
        st.session_state.df_work = st.session_state.df_raw.copy().reset_index(drop=True)
        st.session_state.total = len(st.session_state.df_work)
        st.session_state.row_ptr = 0
        st.session_state.bucket_row_ptr = 0
        st.session_state.file_uploaded = True
        st.session_state.current_category_index = 0
        st.session_state.selected_categories = []
        st.session_state.category_qualifications = []
        st.session_state.qualification_started = False
        st.session_state.category_selection_order = []
        st.session_state.qualifications_by_category = {}
        st.session_state.qualified_categories_by_row = {}
        st.session_state.show_caution_message = False
        st.session_state.confirm_categories = False
        st.session_state.preview_bucket = None
        st.session_state.no_more_records_message = None
        st.success("Excel loaded — start qualifying!")
        safe_rerun()
    except Exception as e:
        st.error(f"Error loading file: {e}")

# ---------- sidebar buckets --------------------------------------------------
st.sidebar.header("👁 Preview Buckets")
bucket_options = [
    "None",
    f"Deleted Records 🗑️ ({len(st.session_state.deleted)})",
    f"To Be Decided ⏳ ({len(st.session_state.to_be_decided)})"
]
selected_bucket_display = st.sidebar.radio(
    "Select a bucket to preview",
    bucket_options,
    index=0,
    key="bucket_selector"
)
# Map display name to internal bucket name
bucket_mapping = {
    "None": None,
    f"Deleted Records 🗑️ ({len(st.session_state.deleted)})": "deleted",
    f"To Be Decided ⏳ ({len(st.session_state.to_be_decided)})": "to_be_decided"
}
st.session_state.preview_bucket = bucket_mapping[selected_bucket_display]

# ---------- preview current row ---------------------------------------------
if st.session_state.file_uploaded and (
    (st.session_state.preview_bucket is None and not st.session_state.df_work.empty) or
    (st.session_state.preview_bucket and not st.session_state[st.session_state.preview_bucket].empty)
):
    if st.session_state.preview_bucket is None:
        i = st.session_state.row_ptr
        source_df = st.session_state.df_work
        total_rows = st.session_state.total
        is_bucket = False
    else:
        i = st.session_state.bucket_row_ptr
        source_df = st.session_state[st.session_state.preview_bucket]
        total_rows = len(st.session_state[st.session_state.preview_bucket])
        is_bucket = True

    # Ensure row is a pandas Series
    row = source_df.iloc[i]

    st.header("Row-by-Row Preview")
    st.markdown(f"### Row {i+1} / {total_rows}")

    st.dataframe(pd.DataFrame(row).T, hide_index=True, use_container_width=True)

    if "URL" in row and pd.notna(row["URL"]):
        st.markdown(f"[**Open Article ↗**]({row['URL']})")

    # Navigation buttons for bucket preview
    if is_bucket:
        col_nav1, col_nav2 = st.columns(2)
        with col_nav1:
            if st.button("Previous ⬅️", key=f"prev_{i}_{st.session_state.preview_bucket}", disabled=(i == 0)):
                st.session_state.bucket_row_ptr = max(0, i - 1)
                st.session_state.no_more_records_message = None
                safe_rerun()
        with col_nav2:
            if st.button("Next ➡️", key=f"next_{i}_{st.session_state.preview_bucket}", disabled=(i >= total_rows - 1)):
                st.session_state.bucket_row_ptr = min(total_rows - 1, i + 1)
                st.session_state.no_more_records_message = None
                safe_rerun()

    # ---------- save_and_advance function ------------------------------------
    def save_and_advance(advance_to_next_row: bool):
        # Save qualifications for the current category
        if st.session_state.current_category_index < len(st.session_state.category_selection_order):
            current_category = st.session_state.category_selection_order[st.session_state.current_category_index]
            qual = st.session_state.qualifications_by_category.get(i, {}).get(current_category, {})
            if qual:
                # Ensure Prominence is a list
                if "Prominence" in qual and (qual["Prominence"] is None or not isinstance(qual["Prominence"], list)):
                    qual["Prominence"] = []
                # Convert Prominence list to a string for storage
                qual_copy = qual.copy()
                qual_copy["Prominence"] = ", ".join(qual["Prominence"]) if qual["Prominence"] else None
                # Use the actual row (pandas Series)
                ann = row.to_frame().T.assign(**qual_copy)
                # Check if Dominance, Prominence, Page, or Tonality are None or empty
                is_partial = (
                    pd.isna(ann["Dominance"]).iloc[0] or
                    (not qual["Prominence"]) or  # Prominence is an empty list
                    pd.isna(ann["Tonality"]).iloc[0]
                    # Page is valid if 0 or greater, so we don't check it for partial
                )
                bucket = "partial" if is_partial else "qualified"
                st.session_state[bucket] = pd.concat(
                    [st.session_state[bucket], ann], ignore_index=True
                )
                if current_category not in st.session_state.qualified_categories_by_row.get(i, []):
                    st.session_state.qualified_categories_by_row.setdefault(i, []).append(current_category)

        # Update state
        if advance_to_next_row:
            if is_bucket:
                current_bucket = st.session_state.preview_bucket
                total_rows_before_drop = len(st.session_state[current_bucket])
                st.session_state[current_bucket].drop(index=i, inplace=True)
                st.session_state[current_bucket].reset_index(drop=True, inplace=True)
                total_rows_new = len(st.session_state[current_bucket])
                if total_rows_new == 0:
                    st.session_state.preview_bucket = None
                    st.session_state.bucket_row_ptr = 0
                    st.session_state.no_more_records_message = f"No more records in the selected preview bucket ('{current_bucket}')."
                else:
                    st.session_state.bucket_row_ptr = min(i + 1, total_rows_new - 1) if total_rows_new > 0 else 0
                    st.session_state.no_more_records_message = None
            else:
                st.session_state.df_work.drop(index=i, inplace=True)
                st.session_state.df_work.reset_index(drop=True, inplace=True)
                st.session_state.total = len(st.session_state.df_work)
                st.session_state.no_more_records_message = None

        if (not is_bucket and st.session_state.df_work.empty) or (is_bucket and st.session_state[st.session_state.preview_bucket].empty):
            st.session_state.row_ptr = 0
            st.session_state.bucket_row_ptr = 0
            st.session_state.file_uploaded = False if not is_bucket else st.session_state.file_uploaded
            st.session_state.selected_categories = []
            st.session_state.category_selection_order = []
            st.session_state.qualifications_by_category = {}
            st.session_state.qualified_categories_by_row = {}
            st.session_state.show_caution_message = False
            st.session_state.confirm_categories = False
            st.session_state.current_category_index = 0
            if is_bucket and st.session_state.no_more_records_message is None:
                st.session_state.preview_bucket = None
        else:
            if advance_to_next_row:
                if is_bucket:
                    # Already updated bucket_row_ptr
                    pass
                else:
                    st.session_state.row_ptr = min(i + 1, len(st.session_state.df_work) - 1) if len(st.session_state.df_work) > 0 else 0
                    qualified_categories = st.session_state.qualified_categories_by_row.get(st.session_state.row_ptr, [])
                    st.session_state.selected_categories = qualified_categories.copy()
                    st.session_state.category_selection_order = qualified_categories.copy()
                    st.session_state.qualifications_by_category[st.session_state.row_ptr] = {}
                    st.session_state.qualified_categories_by_row[st.session_state.row_ptr] = qualified_categories
                    st.session_state.show_caution_message = False
                    st.session_state.confirm_categories = False
                    st.session_state.current_category_index = 0
            else:
                st.session_state.current_category_index += 1
                st.session_state.show_caution_message = (
                    st.session_state.current_category_index >= len(st.session_state.category_selection_order)
                )

        safe_rerun()

    # ---------- save_category_changes function --------------------------------
    def save_category_changes(category: str, q: dict):
        if "Prominence" not in q or q["Prominence"] is None:
            q["Prominence"] = []
        st.session_state.qualifications_by_category.setdefault(i, {})[category] = q
        qual = q.copy()
        qual["Prominence"] = ", ".join(qual["Prominence"]) if qual["Prominence"] else None
        ann = row.to_frame().T.assign(**qual)
        is_partial = (
            pd.isna(ann["Dominance"]).iloc[0] or
            (not q["Prominence"]) or
            pd.isna(ann["Tonality"]).iloc[0]
        )
        bucket = "partial" if is_partial else "qualified"
        for b in ["qualified", "partial"]:
            if not st.session_state[b].empty:
                st.session_state[b] = st.session_state[b][
                    ~((st.session_state[b]["Category"] == category) & (st.session_state[b].index == i))
                ]
        st.session_state[bucket] = pd.concat(
            [st.session_state[bucket], ann], ignore_index=True
        )
        if is_bucket:
            current_bucket = st.session_state.preview_bucket
            total_rows_before_drop = len(st.session_state[current_bucket])
            st.session_state[current_bucket].drop(index=i, inplace=True)
            st.session_state[current_bucket].reset_index(drop=True, inplace=True)
            total_rows_new = len(st.session_state[current_bucket])
            if total_rows_new == 0:
                st.session_state.preview_bucket = None
                st.session_state.bucket_row_ptr = 0
                st.session_state.no_more_records_message = f"No more records in the selected preview bucket ('{current_bucket}')."
            else:
                st.session_state.bucket_row_ptr = min(i + 1, total_rows_new - 1) if total_rows_new > 0 else 0
                st.session_state.no_more_records_message = None
        safe_rerun()

    # Initialize qualifications and qualified categories for this row
    if i not in st.session_state.qualifications_by_category:
        st.session_state.qualifications_by_category[i] = {}
    if i not in st.session_state.qualified_categories_by_row:
        st.session_state.qualified_categories_by_row[i] = []

    # Initialize selected_categories and category_selection_order
    qualified_categories = st.session_state.qualified_categories_by_row.get(i, [])
    if not st.session_state.confirm_categories:
        if not st.session_state.selected_categories:
            st.session_state.selected_categories = qualified_categories.copy()
        if not st.session_state.category_selection_order:
            st.session_state.category_selection_order = qualified_categories.copy()

    # Single-column layout: Select Categories and Qualify Categories stacked vertically
    st.markdown("#### Select Categories")
    if qualified_categories:
        st.info(
            f"**Note**: The following categories have already been qualified for this row: "
            f"{', '.join(qualified_categories)}."
        )
    else:
        st.info("**Note**: No categories have been qualified for this row yet.")

    # Display predefined categories in a 4-column grid
    predefined_categories = []
    previous_selected = st.session_state.selected_categories.copy()
    categories = bank["Category"]
    num_cols = 4
    num_rows = (len(categories) + num_cols - 1) // num_cols  # Ceiling division

    for row_idx in range(num_rows):
        cols = st.columns(num_cols)
        for col_idx, col in enumerate(cols):
            cat_idx = row_idx * num_cols + col_idx
            if cat_idx < len(categories):
                cat = categories[cat_idx]
                with col:
                    default_value = (cat in previous_selected) or (cat in qualified_categories)
                    selected = st.checkbox(cat, key=f"cat_{cat}_{i}", value=default_value, label_visibility="visible")
                    if selected and cat not in st.session_state.selected_categories:
                        st.session_state.selected_categories.append(cat)
                        if cat not in st.session_state.category_selection_order:
                            st.session_state.category_selection_order.append(cat)
                        if st.session_state.show_caution_message:
                            st.session_state.show_caution_message = False
                            safe_rerun()
                    elif not selected and cat in st.session_state.selected_categories:
                        st.session_state.selected_categories.remove(cat)
                        if cat in st.session_state.category_selection_order:
                            st.session_state.category_selection_order.remove(cat)
                    if selected:
                        predefined_categories.append(cat)
                # Place "Add custom category" and "Select saved custom categories" in the same row as "Work Environment"
                if cat_idx == len(categories) - 1:  # When rendering "Work Environment"
                    with cols[1]:  # Second column in the same row
                        new_category = st.text_input(
                            "Add custom category",
                            key=f"add_category_{i}",
                            label_visibility="collapsed",
                            placeholder="Type custom category and press Enter"
                        )
                    with cols[2]:  # Third column in the same row
                        st.markdown("**Select saved custom categories:**")
                        previous_saved_selected = [cat for cat in st.session_state.selected_categories if cat in st.session_state.saved_user_categories]
                        default_saved_selected = list(
                            set(previous_saved_selected + [cat for cat in qualified_categories if cat in st.session_state.saved_user_categories])
                        )
                        saved_selected_categories = st.multiselect(
                            "",
                            st.session_state.saved_user_categories,
                            default=default_saved_selected,
                            key=f"saved_categories_multiselect_{i}",
                            label_visibility="collapsed"
                        )

    # Handle logic for adding and removing categories
    if new_category and new_category not in st.session_state.saved_user_categories:
        st.session_state.saved_user_categories.append(new_category)
        bank["SavedUserCategories"] = st.session_state.saved_user_categories
        save_bank(bank)

    added_categories = [cat for cat in saved_selected_categories if cat not in previous_saved_selected]
    removed_categories = [cat for cat in previous_saved_selected if cat not in saved_selected_categories]
    for cat in removed_categories:
        if cat in st.session_state.selected_categories:
            st.session_state.selected_categories.remove(cat)
        if cat in st.session_state.category_selection_order:
            st.session_state.category_selection_order.remove(cat)
    for cat in added_categories:
        if cat not in st.session_state.selected_categories:
            st.session_state.selected_categories.append(cat)
        if cat not in st.session_state.category_selection_order:
            st.session_state.category_selection_order.append(cat)
        if st.session_state.show_caution_message:
            st.session_state.show_caution_message = False
            safe_rerun()

    categories = predefined_categories + saved_selected_categories
    if new_category and new_category not in categories:
        categories.append(new_category)
        if new_category not in st.session_state.selected_categories:
            st.session_state.selected_categories.append(new_category)
        if new_category not in st.session_state.category_selection_order:
            st.session_state.category_selection_order.append(new_category)
        if st.session_state.show_caution_message:
            st.session_state.show_caution_message = False
            safe_rerun()

    st.session_state.selected_categories = categories

    # Confirm Categories button (immediately below the row)
    if st.button("Confirm Categories ✅", key=f"confirm_categories_{i}", use_container_width=True):
        if not st.session_state.selected_categories:
            st.warning("Please select at least one category before confirming.")
        else:
            st.session_state.confirm_categories = True
            unqualified_categories = [cat for cat in st.session_state.category_selection_order if cat not in qualified_categories]
            st.session_state.current_category_index = (
                st.session_state.category_selection_order.index(unqualified_categories[0])
                if unqualified_categories else 0
            )
            st.session_state.show_caution_message = False
            st.session_state.no_more_records_message = None
            safe_rerun()

    st.divider()

    # Qualify for the Selected Category section
    if st.session_state.confirm_categories and st.session_state.selected_categories and st.session_state.category_selection_order:
        if st.session_state.current_category_index < len(st.session_state.category_selection_order):
            current_category = st.session_state.category_selection_order[st.session_state.current_category_index]
            total_categories = len(st.session_state.category_selection_order)
            current_category_num = st.session_state.current_category_index + 1
            st.markdown(f"#### Qualify for '{current_category}' Category ({current_category_num}/{total_categories})")

            current_qualifications = st.session_state.qualifications_by_category.get(i, {}).get(current_category, {})
            q = {}
            q["Category"] = current_category

            # Arrange fields in a single row
            col_d, col_p, col_s, col_pg, col_t = st.columns([1, 1, 1, 1, 1])

            with col_d:
                st.markdown("**Dominance**")
                default_dominance = current_qualifications.get("Dominance")
                q["Dominance"] = st.radio(
                    "",
                    bank["Dominance"],
                    index=bank["Dominance"].index(default_dominance) if default_dominance in bank["Dominance"] else None,
                    key=f"sel_dominance_{i}_{st.session_state.current_category_index}",
                    label_visibility="collapsed"
                )

            with col_p:
                st.markdown("**Prominence**")
                default_prominence = current_qualifications.get("Prominence", [])
                if default_prominence is None:
                    default_prominence = []
                prominence_selections = []
                for option in bank["Prominence"]:
                    selected = option in default_prominence
                    if st.checkbox(
                        option,
                        value=selected,
                        key=f"prominence_{option}_{i}_{current_category}",
                        label_visibility="visible"
                    ):
                        prominence_selections.append(option)
                q["Prominence"] = prominence_selections

            with col_s:
                st.markdown("**Spokesperson**")
                default_spokesperson = current_qualifications.get("Spokesperson")
                q["Spokesperson"] = st.radio(
                    "",
                    bank["Spokesperson"],
                    index=bank["Spokesperson"].index(default_spokesperson) if default_spokesperson in bank["Spokesperson"] else None,
                    key=f"sel_spokesperson_{i}_{st.session_state.current_category_index}",
                    label_visibility="collapsed"
                )

            with col_pg:
                st.markdown("**Page**")
                default_page = current_qualifications.get("Page", 0)
                default_page = 0 if default_page is None else default_page
                q["Page"] = st.number_input(
                    "",
                    min_value=0,
                    step=1,
                    value=default_page,
                    key=f"page_{i}_{current_category}",
                    label_visibility="collapsed"
                )

            with col_t:
                st.markdown("**Tonality**")
                default_tonality = current_qualifications.get("Tonality")
                q["Tonality"] = st.radio(
                    "",
                    bank["Tonality"],
                    index=bank["Tonality"].index(default_tonality) if default_tonality in bank["Tonality"] else None,
                    key=f"sel_tonality_{i}_{st.session_state.current_category_index}",
                    label_visibility="collapsed"
                )

            # Spokesperson Name with Designation (below if Spokesperson is selected)
            if q["Spokesperson"]:
                st.markdown("**Spokesperson Name with Designation**")
                default_spokesperson_name = current_qualifications.get("Spokesperson Name with Designation", "")
                q["Spokesperson Name with Designation"] = st.text_input(
                    "",
                    value=default_spokesperson_name,
                    key=f"spokesperson_name_{i}_{current_category}",
                    label_visibility="collapsed"
                )
            else:
                q["Spokesperson Name with Designation"] = None

            # Change: Conditionally show "Save & Qualify Further" or "Save & Review" button
            # Check if this is the last category to qualify
            is_last_category = (st.session_state.current_category_index + 1) == len(st.session_state.category_selection_order)
            if is_last_category:
                if st.button("Save & Review 📋", key=f"save_review_{i}"):
                    missing_fields = []
                    if q["Dominance"] is None:
                        missing_fields.append("Dominance")
                    if q["Tonality"] is None:
                        missing_fields.append("Tonality")
                    if missing_fields:
                        st.warning(f"Please select values for the following mandatory fields: {', '.join(missing_fields)}.")
                    else:
                        # Save and move to review by setting show_caution_message
                        st.session_state.current_category_index += 1
                        st.session_state.show_caution_message = True
                        safe_rerun()
            else:
                if st.button("Save & Qualify Further 💾", key=f"save_qualify_{i}"):
                    missing_fields = []
                    if q["Dominance"] is None:
                        missing_fields.append("Dominance")
                    if q["Tonality"] is None:
                        missing_fields.append("Tonality")
                    if missing_fields:
                        st.warning(f"Please select values for the following mandatory fields: {', '.join(missing_fields)}.")
                    else:
                        save_and_advance(False)

            st.session_state.qualifications_by_category.setdefault(i, {})[current_category] = q
        else:
            st.markdown("#### Review Qualified Categories")
            review_category = st.selectbox(
                "Select a category to review qualifications",
                options=st.session_state.category_selection_order,
                key=f"review_category_{i}"
            )
            if review_category:
                current_qualifications = st.session_state.qualifications_by_category.get(i, {}).get(review_category, {})
                q = {}
                q["Category"] = review_category

                col_d, col_p, col_s, col_pg, col_t = st.columns([1, 1, 1, 1, 1])

                with col_d:
                    st.markdown("**Dominance**")
                    default_dominance = current_qualifications.get("Dominance")
                    q["Dominance"] = st.selectbox(
                        "",
                        ["— select —"] + bank["Dominance"],
                        index=0 if default_dominance is None else bank["Dominance"].index(default_dominance) + 1,
                        key=f"review_dominance_{review_category}_{i}",
                        label_visibility="collapsed"
                    )
                    if q["Dominance"] == "— select —":
                        q["Dominance"] = None

                with col_p:
                    st.markdown("**Prominence**")
                    default_prominence = current_qualifications.get("Prominence", [])
                    if default_prominence is None:
                        default_prominence = []
                    q["Prominence"] = st.multiselect(
                        "",
                        bank["Prominence"],
                        default=default_prominence,
                        key=f"review_prominence_{i}_{review_category}",
                        label_visibility="collapsed"
                    )

                with col_s:
                    st.markdown("**Spokesperson**")
                    default_spokesperson = current_qualifications.get("Spokesperson")
                    q["Spokesperson"] = st.selectbox(
                        "",
                        ["— select —"] + bank["Spokesperson"],
                        index=0 if default_spokesperson is None else bank["Spokesperson"].index(default_spokesperson) + 1,
                        key=f"review_spokesperson_{review_category}_{i}",
                        label_visibility="collapsed"
                    )
                    if q["Spokesperson"] == "— select —":
                        q["Spokesperson"] = None

                with col_pg:
                    st.markdown("**Page**")
                    default_page = current_qualifications.get("Page", 0)
                    default_page = 0 if default_page is None else default_page
                    q["Page"] = st.number_input(
                        "",
                        min_value=0,
                        step=1,
                        value=default_page,
                        key=f"review_page_{i}_{review_category}",
                        label_visibility="collapsed"
                    )

                with col_t:
                    st.markdown("**Tonality**")
                    default_tonality = current_qualifications.get("Tonality")
                    q["Tonality"] = st.selectbox(
                        "",
                        ["— select —"] + bank["Tonality"],
                        index=0 if default_tonality is None else bank["Tonality"].index(default_tonality) + 1,
                        key=f"review_tonality_{review_category}_{i}",
                        label_visibility="collapsed"
                    )
                    if q["Tonality"] == "— select —":
                        q["Tonality"] = None

                if q["Spokesperson"]:
                    st.markdown("**Spokesperson Name with Designation**")
                    default_spokesperson_name = current_qualifications.get("Spokesperson Name with Designation", "")
                    q["Spokesperson Name with Designation"] = st.text_input(
                        "",
                        value=default_spokesperson_name,
                        key=f"review_spokesperson_name_{i}_{review_category}",
                        label_visibility="collapsed"
                    )
                else:
                    q["Spokesperson Name with Designation"] = None

                if st.button("Save Changes for this Category 💾", key=f"save_review_{i}_{review_category}"):
                    missing_fields = []
                    if q["Dominance"] is None:
                        missing_fields.append("Dominance")
                    if q["Tonality"] is None:
                        missing_fields.append("Tonality")
                    if missing_fields:
                        st.warning(f"Please select values for the following mandatory fields: {', '.join(missing_fields)}.")
                    else:
                        save_category_changes(review_category, q)
            else:
                st.info("All selected categories have been qualified. Please select a category to review or click 'Save & Next' to proceed.")
    else:
        st.info("Please select and confirm categories above to start qualifying.")

    if not st.session_state.selected_categories and st.session_state.show_caution_message:
        all_categories = bank["Category"] + st.session_state.saved_user_categories
        qualified_categories = st.session_state.qualified_categories_by_row.get(i, [])
        non_qualified_categories = [cat for cat in all_categories if cat not in qualified_categories]
        if non_qualified_categories:
            st.warning(
                "Please select categories from non-selected categories or click on Save & Next to proceed ahead."
            )
        else:
            st.warning(
                "All categories have been qualified for this row. Click on Save & Next to proceed ahead."
            )

    if not st.session_state.selected_categories:
        c1, c2 = st.columns(2)
        to_be_decided = c1.button("To Be Decided ⏳", key=f"to_be_decided_{i}", use_container_width=True)
        delete = c2.button("Delete 🗑️", key=f"del_{i}", use_container_width=True)

        def advance(delete_row: bool, bucket: str | None = None, payload=None):
            if delete_row:
                if is_bucket:
                    current_bucket = st.session_state.preview_bucket
                    total_rows_before_drop = len(st.session_state[current_bucket])
                    st.session_state[current_bucket].drop(index=i, inplace=True)
                    st.session_state[current_bucket].reset_index(drop=True, inplace=True)
                    total_rows_new = len(st.session_state[current_bucket])
                    if total_rows_new == 0:
                        st.session_state.preview_bucket = None
                        st.session_state.bucket_row_ptr = 0
                        st.session_state.no_more_records_message = f"No more records in the selected preview bucket ('{current_bucket}')."
                    else:
                        st.session_state.bucket_row_ptr = min(i + 1, total_rows_new - 1) if total_rows_new > 0 else 0
                        st.session_state.no_more_records_message = None
                else:
                    st.session_state.df_work.drop(index=i, inplace=True)
                    st.session_state.df_work.reset_index(drop=True, inplace=True)
                    st.session_state.total = len(st.session_state.df_work)
                    if st.session_state.total == 0:
                        st.session_state.row_ptr = 0
                    else:
                        st.session_state.row_ptr = min(i + 1, st.session_state.total - 1) if st.session_state.total > 0 else 0
                    st.session_state.no_more_records_message = None

            if bucket and payload is not None:
                st.session_state[bucket] = pd.concat(
                    [st.session_state[bucket], payload], ignore_index=True
                )

            if (not is_bucket and st.session_state.df_work.empty) or (is_bucket and st.session_state[st.session_state.preview_bucket].empty):
                st.session_state.row_ptr = 0
                st.session_state.bucket_row_ptr = 0
                st.session_state.file_uploaded = False if not is_bucket else st.session_state.file_uploaded
                st.session_state.selected_categories = []
                st.session_state.category_selection_order = []
                st.session_state.qualifications_by_category = {}
                st.session_state.qualified_categories_by_row = {}
                st.session_state.show_caution_message = False
                st.session_state.confirm_categories = False
                st.session_state.current_category_index = 0
                if is_bucket and st.session_state.no_more_records_message is None:
                    st.session_state.preview_bucket = None
            else:
                new_i = st.session_state.bucket_row_ptr if is_bucket else st.session_state.row_ptr
                if new_i not in st.session_state.qualifications_by_category:
                    st.session_state.qualifications_by_category[new_i] = {}
                if new_i not in st.session_state.qualified_categories_by_row:
                    st.session_state.qualified_categories_by_row[new_i] = []
                st.session_state.selected_categories = st.session_state.qualified_categories_by_row[new_i].copy()
                st.session_state.category_selection_order = st.session_state.qualified_categories_by_row[new_i].copy()
                st.session_state.show_caution_message = False
                st.session_state.confirm_categories = False
                st.session_state.current_category_index = 0

            safe_rerun()

        if to_be_decided:
            advance(True, "to_be_decided", row.to_frame().T)
        elif delete:
            advance(True, "deleted", row.to_frame().T)

else:
    if not st.session_state.file_uploaded:
        st.info("Please upload an Excel file to start qualifying.")
    else:
        if st.session_state.no_more_records_message:
            st.warning(st.session_state.no_more_records_message)
        else:
            st.info("No rows to qualify in the selected dataset. Please upload another Excel file or select a different bucket.")

# Save & Next button - Hide if on the last record in a preview bucket
if (st.session_state.file_uploaded and
    ((st.session_state.preview_bucket is None and not st.session_state.df_work.empty) or
     (st.session_state.preview_bucket and not st.session_state[st.session_state.preview_bucket].empty)) and
    st.session_state.confirm_categories and
    not (is_bucket and i == total_rows - 1)):
    st.button("Save & Next ➡️", key=f"save_next_{i}", use_container_width=True, on_click=lambda: save_and_advance(True))

# Download Qualified Data button
if not st.session_state.qualified.empty or not st.session_state.partial.empty:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as wr:
        all_qualified = pd.concat(
            [st.session_state.qualified, st.session_state.partial], ignore_index=True
        )
        all_qualified.to_excel(wr, index=False, sheet_name="All Qualified Data")
    buf.seek(0)
    st.download_button(
        "Download Qualified Data (Excel)",
        buf.read(),
        file_name="qualified_news_items.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

