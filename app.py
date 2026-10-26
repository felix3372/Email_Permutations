"""
Email Permutation Generator (standalone, single file)

Run:   pip install streamlit pandas openpyxl
       streamlit run app.py

Input : First Name, Last Name, Email, Domain
Output: First Name, Last Name, Email, Domain, New Email
        (one row per permutation; Email is kept only as the lead reference)
"""
import io
import re
import unicodedata

import pandas as pd
import streamlit as st

# =============================================================================
# EMAIL PERMUTATION ENGINE (from email_utils.py, logic unchanged)
# =============================================================================

# ✅ Centralized ignore suffixes - Single source of truth
IGNORE_SUFFIXES = {
    # Titles
    "dr", "dr.", "mr", "mr.", "ms", "ms.", "mrs", "mrs.", "prof", "prof.",
    "sir", "hon", "hon.", "rev", "rev.", "fr", "fr.",
    # Military ranks
    "capt", "capt.", "captain", "col", "col.", "colonel", "maj", "maj.", 
    "lt", "lt.", "lieutenant", "sgt", "sgt.", "sergeant", "pvt", "pvt.", "private",
    "general", "admiral", "commander", "corporal", "cpl", "cpl.",
    # Religious titles
    "minister", "archbishop", "rabbi", "brother", "sister", "father", "mother",
    # Academic/Professional titles
    "chancellor", "principal", "director", "president", "ceo", 
    "cfo", "cto", "cmo", "vp", "svp", "evp", "chair", "chairman", "chairwoman",
    # Status indicators
    "ret", "ret.", "retired", "emeritus",
    # Generational suffixes
    "jr", "jr.", "sr", "sr.", "junior", "senior", "ii", "iii", "iv", "v", "vi", "2nd", "3rd", "4th", "5th",
    # Degrees & Certifications
    "mba", "phd", "msc", "ms", "bsc", "ba", "ma", "mtech", "btech", "jd", "ed.d",
    "cfa", "cpa", "acca", "ca", "cma", "cfp", "frms", "pmp", "sixsigma", "six",
    "sigma", "lssbb", "lssgb", "lean", "blackbelt", "rn", "md", "mbbs", "dds", "dmd",
    "do", "bpharm", "mph", "shrmp", "phr", "gphr", "sphr", "ccmp", "cpc", "cpcc",
    "aws", "gcp", "ccna", "mcsa", "ocp", "mcp", "comptia", "azure", "esq", "llb", "llm",
    "itil", "itl", "iti", "cissp", "ciso", "cism", "cris", "bms", "flmi", "flmi.",
    "hia", "hia.", "mhp", "mhp.", "acs", "acs.", "cpcu", "cpcu.", "csm", "csm.",
    "gaicd", "gaicd.", "cmaicd", "cmaicd.", "fcipd", "fcipd.", "cipd", "cipd.",
    "fcpd", "fcpd.", "fcp", "fcp.", "fca", "fca.", "fcma", "fcma.", "fcpa", "fcpa.",
    "fbcs", "fbcs."
}


def extract_root_domain(domain):
    """
    Extracts the domain 'root' (registered domain only) using tldextract.
    E.g., maaden.com.sa → maaden
    """
    if not domain:
        return ""
    import tldextract  # lazy import: only needed if this helper is used
    ext = tldextract.extract(domain)
    return ext.domain if ext.domain else domain.lower()


def strip_unicode(text):
    replacements = {
        "ü": "ue", "ö": "oe", "ä": "ae", "ß": "ss",
        "Ü": "ue", "Ö": "oe", "Ä": "ae", "ẞ": "ss"
    }
    for char, rep in replacements.items():
        text = text.replace(char, rep)
    return unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii')


def extract_primary_first_name(name):
    m = re.match(r"([^\(]+)", name)
    if m:
        return m.group(1).strip()
    return name.strip()


def tokenize_name(name, allowed_separators=r"[\s\-\.'_]+"):
    """
    Tokenize any name by common separators (spaces, hyphens, periods, apostrophes, underscores)
    Returns list of cleaned tokens
    """
    if not name:
        return []
    
    # Split by separators and clean each token
    raw_tokens = re.split(allowed_separators, name.strip())
    tokens = []
    
    for token in raw_tokens:
        cleaned = strip_unicode(re.sub(r"[^a-z0-9._-]", "", token.lower()))
        if cleaned:  # Only add non-empty tokens
            tokens.append(cleaned)
    
    return tokens


def tokenize_last_name(last_name, allowed_separators=r"[\s\-\.'_]+"):
    """
    Wrapper for backward compatibility - just calls tokenize_name
    """
    return tokenize_name(last_name, allowed_separators)


def restructure_compound_names(first_name, last_name):
    """
    Smart name restructuring for compound first names that filters out titles/suffixes.
    If first name has multiple tokens, keep only the first valid token as first name
    and move the rest to the beginning of last name (excluding titles/suffixes).
    
    Examples:
    - "Jesus Ortiz", "Lopez" → "Jesus", "Ortiz Lopez"
    - "Dr. Lisa", "Ali" → "Lisa", "Ali" (Dr. filtered out)
    - "John Michael", "Smith Jones" → "John", "Michael Smith Jones"
    - "Prof. Maria Elena", "Garcia" → "Maria", "Elena Garcia" (Prof. filtered out)
    """
    if not first_name or not last_name:
        return first_name or "", last_name or ""
    
    first_tokens = first_name.strip().split()
    last_tokens = last_name.strip().split()
    
    if len(first_tokens) > 1:
        # Filter out titles/suffixes from first name tokens using centralized list
        valid_first_tokens = [
            token for token in first_tokens 
            if token.lower().rstrip('.') not in IGNORE_SUFFIXES
        ]
        
        if valid_first_tokens:
            # Keep only first valid token as first name
            restructured_first = valid_first_tokens[0]
            # Move remaining valid tokens to beginning of last name (exclude titles)
            extra_tokens = valid_first_tokens[1:]
            restructured_last = " ".join(extra_tokens + last_tokens) if extra_tokens else " ".join(last_tokens)
            return restructured_first, restructured_last
        else:
            # All tokens were filtered out, fallback to original
            return first_name.strip(), last_name.strip()
    
    return first_name.strip(), last_name.strip()


def generate_email_permutations(first_name, last_name, domain, 
                               token_min_len=1,  # ✅ Default 1 for comprehensive suppression
                               row_budget=100,   # ✅ Increased default for suppression coverage
                               enable_token_mode=True):
    """
    Enhanced email permutation generator with token-based last name approach
    and smart compound name restructuring.
    
    DEFAULT CONFIGURATION: Optimized for suppression checks (comprehensive pattern generation)
    
    Args:
        first_name: First name
        last_name: Last name  
        domain: Email domain
        token_min_len: Minimum length for last name tokens (default 1 for suppression)
        row_budget: Maximum extra candidates from token mode (default 100 for suppression)
        enable_token_mode: Whether to enable token-based enhancement (default True)
    
    Returns:
        Set of email permutations
    """
    
    # Store original names for additional pattern generation
    original_first_name = first_name
    original_last_name = last_name
    
    # ✅ Smart name restructuring for compound first names
    restructured_first, restructured_last = restructure_compound_names(first_name, last_name)
    
    # Generate patterns for BOTH restructured and original names
    all_permutations = set()
    
    # 1. Generate patterns with restructured names (existing logic)
    restructured_patterns = _generate_patterns_core(
        restructured_first, restructured_last, domain, 
        token_min_len, row_budget, enable_token_mode, 
        original_first_name, original_last_name  # Pass original names for compound patterns
    )
    all_permutations.update(restructured_patterns)
    
    # 2. Generate patterns with original names (for cultural naming patterns)
    # Only if restructuring actually changed something
    if (restructured_first != original_first_name or restructured_last != original_last_name):
        original_patterns = _generate_patterns_core(
            original_first_name, original_last_name, domain,
            token_min_len, row_budget, enable_token_mode,
            original_first_name, original_last_name  # Pass same as originals
        )
        all_permutations.update(original_patterns)
    
    return all_permutations


def _generate_patterns_core(first_name, last_name, domain, token_min_len, row_budget, 
                           enable_token_mode, original_first_name=None, original_last_name=None):
    """
    Core pattern generation logic (extracted to avoid duplication)
    Now includes support for compound first name tokens
    """
    
    # Handle primary first name (before any parentheses)
    raw_first = (first_name or "").strip().lower()
    first_cleaned = strip_unicode(extract_primary_first_name(raw_first))
    first_parts = re.split(r"\s+", first_cleaned)
    first_parts = [re.sub(r"[^a-z0-9._-]", "", fp) for fp in first_parts if fp]

    domain = (domain or "").strip().lower()

    # ⛏️ Split raw last name before cleaning to preserve word boundaries
    raw_last_parts = (last_name or "").strip().lower().split()
    last_parts = [strip_unicode(re.sub(r"[^a-z0-9._-]", "", p)) for p in raw_last_parts]

    if not first_parts or not last_parts or not domain:
        return set()

    # Filter out ignore suffixes using centralized list
    last_parts = [part for part in last_parts if part.lower() not in IGNORE_SUFFIXES]

    if not last_parts:
        return set()

    l_full = ''.join(last_parts)
    f_full = ''.join(first_parts)
    f_initial = first_parts[0][0] if first_parts else ""
    l_initial = last_parts[0][0] if last_parts and last_parts[0] else ""
    last_word = last_parts[-1] if last_parts else ""

    permutations = set()

    # Standard f+l combinations (each first name part x each last name part)
    for f in first_parts:
        for l in last_parts:
            permutations.update([
                f"{f}{l}@{domain}",
                f"{f}.{l}@{domain}",
                f"{f}_{l}@{domain}",
                f"{f}-{l}@{domain}",
                f"{l}{f}@{domain}",
                f"{l}.{f}@{domain}",
                f"{l}_{f}@{domain}",
                f"{l}-{f}@{domain}",
            ])

    # Full combos and swapped
    permutations.update([
        f"{f_full}{l_full}@{domain}",
        f"{f_full}.{l_full}@{domain}",
        f"{f_full}_{l_full}@{domain}",
        f"{f_full}-{l_full}@{domain}",
        f"{l_full}{f_full}@{domain}",
        f"{l_full}.{f_full}@{domain}",
        f"{l_full}_{f_full}@{domain}",
        f"{l_full}-{f_full}@{domain}",
        f"{f_full}@{domain}",
        f"{l_full}@{domain}",
        f"{f_full}{l_initial}@{domain}" if l_initial else "",
        f"{l_full}{f_initial}@{domain}" if f_initial else "",
        f"{f_full}.{l_initial}@{domain}" if l_initial else "",
        f"{l_full}.{f_initial}@{domain}" if f_initial else "",
        f"{f_initial}.{l_full}@{domain}" if f_initial else "",
        f"{l_initial}.{f_full}@{domain}" if l_initial else "",
        f"{f_initial}_{l_full}@{domain}" if f_initial else "",
        f"{l_initial}_{f_full}@{domain}" if l_initial else "",
        f"{f_initial}-{l_full}@{domain}" if f_initial else "",
        f"{l_initial}-{f_full}@{domain}" if l_initial else "",
        f"{f_initial}{l_full}@{domain}" if f_initial else "",
        f"{l_full}{f_initial}@{domain}" if f_initial else "",
    ])

    # Add first_lastinitial, first.lastinitial, first-lastinitial for each last name part
    if f_full and last_parts:
        for lp in last_parts:
            if lp:
                lp_initial = lp[0]
                permutations.update([
                    f"{f_full}{lp_initial}@{domain}",
                    f"{f_full}_{lp_initial}@{domain}",
                    f"{f_full}.{lp_initial}@{domain}",
                    f"{f_full}-{lp_initial}@{domain}",
                ])

    # Add last_firstinitial patterns (williams_j, williams.j, williams-j, williamsj)
    if l_full and f_initial:
        permutations.update([
            f"{l_full}_{f_initial}@{domain}",
            f"{l_full}.{f_initial}@{domain}",
            f"{l_full}-{f_initial}@{domain}",
            f"{l_full}{f_initial}@{domain}",
        ])

    # Add first initial + last, last initial + first (e.g. smac@, mstan@)
    if f_initial and l_full:
        permutations.add(f"{f_initial}{l_full}@{domain}")
    if l_initial and f_full:
        permutations.add(f"{l_initial}{f_full}@{domain}")
        permutations.update([
            f"{l_initial}_{f_full}@{domain}",
            f"{l_initial}.{f_full}@{domain}",
            f"{l_initial}-{f_full}@{domain}",
        ])

    # Final fallback formats using last word of last name
    if last_word:
        if f_full:
            permutations.update([
                f"{f_full}.{last_word}@{domain}",
                f"{f_full}{last_word}@{domain}",
            ])
        if f_initial:
            permutations.update([
                f"{f_initial}.{last_word}@{domain}",
                f"{f_initial}{last_word}@{domain}",
            ])

    # Add first_full + "_" + l_initial (e.g. chanda_c@...)
    if f_full and l_initial:
        permutations.add(f"{f_full}_{l_initial}@{domain}")

    # ✅ Concatenated Multi-Token Patterns for compound last names
    # If last name has multiple parts, also try concatenating them (e.g., "ortiz lopez" → "ortizlopez")
    if len(last_parts) > 1:
        concatenated_last = ''.join(last_parts)  # "ortizlopez"
        permutations.update([
            f"{f_full}.{concatenated_last}@{domain}",
            f"{f_full}{concatenated_last}@{domain}",
            f"{f_full}_{concatenated_last}@{domain}",
            f"{f_full}-{concatenated_last}@{domain}",
            f"{concatenated_last}.{f_full}@{domain}",
            f"{concatenated_last}{f_full}@{domain}",
            f"{concatenated_last}_{f_full}@{domain}",
            f"{concatenated_last}-{f_full}@{domain}",
            f"{concatenated_last}@{domain}",
            f"{f_initial}.{concatenated_last}@{domain}" if f_initial else "",
            f"{f_initial}{concatenated_last}@{domain}" if f_initial else "",
            f"{concatenated_last}{f_initial}@{domain}" if f_initial else "",
            f"{concatenated_last}.{f_initial}@{domain}" if f_initial else "",
            f"{concatenated_last}_{f_initial}@{domain}" if f_initial else "",
            f"{concatenated_last}-{f_initial}@{domain}" if f_initial else "",
        ])

    # ✅ NEW: Handle compound FIRST names (like "Ba Loc")
    # Generate patterns using individual tokens from compound first names
    if len(first_parts) > 1:
        # For compound first names like "Ba Loc", generate patterns like ba.loc@
        for i in range(len(first_parts)):
            for j in range(i+1, len(first_parts)):
                permutations.update([
                    f"{first_parts[i]}.{first_parts[j]}@{domain}",
                    f"{first_parts[i]}_{first_parts[j]}@{domain}",
                    f"{first_parts[i]}-{first_parts[j]}@{domain}",
                    f"{first_parts[i]}{first_parts[j]}@{domain}",
                ])
        
        # Also generate patterns with each first name part combined with last name parts
        for fp in first_parts:
            for lp in last_parts:
                permutations.update([
                    f"{fp}.{lp}@{domain}",
                    f"{fp}_{lp}@{domain}",
                    f"{fp}-{lp}@{domain}",
                    f"{fp}{lp}@{domain}",
                ])

    # If token mode is disabled, return original results
    if not enable_token_mode:
        # Remove any blank permutations
        permutations = {p for p in permutations if p}
        return permutations
    
    # ===== ENHANCED TOKEN-BASED APPROACH =====
    token_permutations = set()
    
    # Tokenize BOTH names for comprehensive coverage
    last_tokens = tokenize_name(last_name)
    first_tokens = tokenize_name(first_name)
    
    # ✅ NEW: Also tokenize original names if provided (for compound name support)
    if original_first_name and original_first_name != first_name:
        original_first_tokens = tokenize_name(original_first_name)
        first_tokens.extend(original_first_tokens)
    
    if original_last_name and original_last_name != last_name:
        original_last_tokens = tokenize_name(original_last_name)
        last_tokens.extend(original_last_tokens)
    
    # Remove duplicates and filter tokens
    all_tokens = list(set(first_tokens + last_tokens))
    
    # Filter tokens by minimum length and ignore suffixes
    valid_tokens = [
        token for token in all_tokens 
        if len(token) >= token_min_len and token.lower() not in IGNORE_SUFFIXES
    ]
    
    # Prioritize longer tokens first (more meaningful)
    valid_tokens.sort(key=len, reverse=True)
    
    # Generate patterns for each token, respecting budget
    candidates_added = 0
    
    for token in valid_tokens:
        if candidates_added >= row_budget:
            break
            
        # Generate all patterns for this token
        token_patterns = set()
        
        # Patterns with each first name part and the token
        for f in first_parts:
            token_patterns.update([
                f"{f}{token}@{domain}",
                f"{f}.{token}@{domain}",
                f"{f}_{token}@{domain}",
                f"{f}-{token}@{domain}",
                f"{token}{f}@{domain}",
                f"{token}.{f}@{domain}",
                f"{token}_{f}@{domain}",
                f"{token}-{f}@{domain}",
            ])
        
        # Full first name with token
        token_patterns.update([
            f"{f_full}{token}@{domain}",
            f"{f_full}.{token}@{domain}",
            f"{f_full}_{token}@{domain}",
            f"{f_full}-{token}@{domain}",
            f"{token}{f_full}@{domain}",
            f"{token}.{f_full}@{domain}",
            f"{token}_{f_full}@{domain}",
            f"{token}-{f_full}@{domain}",
            f"{token}@{domain}",
        ])
        
        # Initial combinations with token
        if f_initial:
            token_patterns.update([
                f"{f_initial}.{token}@{domain}",
                f"{f_initial}{token}@{domain}",
                f"{token}{f_initial}@{domain}",
                f"{token}.{f_initial}@{domain}",
                f"{token}_{f_initial}@{domain}",
                f"{token}-{f_initial}@{domain}",
            ])
        
        # ✅ NEW: Token-to-token patterns (for compound names)
        for other_token in valid_tokens:
            if other_token != token:
                token_patterns.update([
                    f"{token}.{other_token}@{domain}",
                    f"{token}_{other_token}@{domain}",
                    f"{token}-{other_token}@{domain}",
                    f"{token}{other_token}@{domain}",
                ])
        
        # Add new patterns that weren't in original set
        new_patterns = token_patterns - permutations - token_permutations
        
        # Respect budget
        budget_remaining = row_budget - candidates_added
        patterns_to_add = list(new_patterns)[:budget_remaining]
        
        token_permutations.update(patterns_to_add)
        candidates_added += len(patterns_to_add)
    
    # Combine original and token-based permutations
    all_permutations = permutations | token_permutations
    
    # Remove any blank permutations
    all_permutations = {p for p in all_permutations if p}
    
    return all_permutations


def generate_validation_patterns(first_name, last_name, domain):
    """
    Email patterns optimized for validation (stricter filtering).
    
    Uses stricter token filtering to avoid unrealistic patterns like single-letter emails.
    Perfect for email_validator.py where we want realistic, probable email formats.
    
    Args:
        first_name: First name
        last_name: Last name  
        domain: Email domain
    
    Returns:
        Set of realistic email permutations for validation
    """
    return generate_email_permutations(
        first_name, 
        last_name, 
        domain, 
        token_min_len=2,      # Stricter: exclude single-char tokens  
        row_budget=60,        # Reasonable limit for validation
        enable_token_mode=True
    )


# =============================================================================
# STREAMLIT APP
# =============================================================================

st.set_page_config(page_title="Email Permutation Generator", page_icon="✉️", layout="wide")

REQUIRED = ["First Name", "Last Name", "Email", "Domain"]
OUTPUT_COLS = REQUIRED + ["New Email"]

# Accepted header spellings (compared after lowercasing and removing spaces/_/-)
ALIASES = {
    "First Name": {"firstname", "first", "fname", "givenname"},
    "Last Name": {"lastname", "last", "lname", "surname", "familyname"},
    "Email": {"email", "emailaddress", "mail", "workemail"},
    "Domain": {"domain", "companydomain", "website", "emaildomain"},
}


def norm(s: str) -> str:
    return re.sub(r"[\s_\-]+", "", str(s).strip().lower())


def map_columns(df: pd.DataFrame):
    """Return ({required_name: actual_column}, [missing required names])."""
    mapping, missing = {}, []
    by_norm = {norm(c): c for c in df.columns}
    for req, aliases in ALIASES.items():
        hit = next((by_norm[a] for a in aliases if a in by_norm), None)
        if hit is None:
            missing.append(req)
        else:
            mapping[req] = hit
    return mapping, missing


def clean_domain(raw: str) -> str:
    """'https://www.Acme.com/about' -> 'acme.com'. Keeps the full domain incl. TLD."""
    d = (raw or "").strip().lower()
    d = re.sub(r"^[a-z]+://", "", d)
    d = d.split("/")[0].split("?")[0].split("#")[0]
    if "@" in d:  # someone put an email in the domain column
        d = d.split("@")[-1]
    d = re.sub(r"^www\.", "", d)
    return d.strip(" .")


def read_upload(file) -> pd.DataFrame:
    name = file.name.lower()
    if name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(file, dtype=str, keep_default_na=False)
    else:
        df = pd.read_csv(file, dtype=str, keep_default_na=False, encoding_errors="replace")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def build_output(df, mapping, mode, token_min_len, row_budget, token_mode, drop_existing):
    """Expand each lead into one row per permutation. Original columns repeat on every row."""
    rows, skipped = [], []

    for _, r in df.iterrows():
        first = str(r[mapping["First Name"]]).strip()
        last = str(r[mapping["Last Name"]]).strip()
        email = str(r[mapping["Email"]]).strip()
        domain_raw = str(r[mapping["Domain"]]).strip()
        domain = clean_domain(domain_raw)

        if mode == "Validation (stricter, fewer patterns)":
            perms = generate_validation_patterns(first, last, domain)
        else:
            perms = generate_email_permutations(
                first, last, domain,
                token_min_len=token_min_len,
                row_budget=row_budget,
                enable_token_mode=token_mode,
            )

        if drop_existing and email:
            perms = {p for p in perms if p.lower() != email.lower()}

        if not perms:
            skipped.append({"First Name": first, "Last Name": last, "Email": email, "Domain": domain_raw})
            continue

        for p in sorted(perms):
            rows.append((first, last, email, domain_raw, p))

    out = pd.DataFrame(rows, columns=OUTPUT_COLS)
    return out, pd.DataFrame(skipped, columns=REQUIRED)


# ----------------------------------------------------------------------------- UI
st.title("✉️ Email Permutation Generator")
st.caption(
    "Upload leads with **First Name, Last Name, Email, Domain**. You get back the same columns plus "
    "**New Email**: one row per permutation. The original Email stays on every row so you can match "
    "each New Email back to its lead."
)

with st.sidebar:
    st.header("Settings")
    mode = st.radio(
        "Pattern mode",
        ["Comprehensive (suppression)", "Validation (stricter, fewer patterns)"],
        help="Comprehensive = every pattern your email_utils can produce. "
             "Validation = realistic patterns only (min token length 2, budget 60).",
    )
    comprehensive = mode.startswith("Comprehensive")
    token_mode = st.checkbox("Token-based patterns", value=True, disabled=not comprehensive)
    token_min_len = st.number_input("Min token length", 1, 10, 1, disabled=not comprehensive)
    row_budget = st.number_input("Extra token patterns per lead (budget)", 0, 1000, 100, disabled=not comprehensive)
    drop_existing = st.checkbox(
        "Exclude lead's existing Email from New Email", value=False,
        help="Removes a permutation if it is identical to the Email already on that lead's row.",
    )

file = st.file_uploader("Upload leads file", type=["csv", "xlsx", "xls"])

if file:
    try:
        df = read_upload(file)
    except Exception as e:
        st.error(f"Couldn't read that file: {e}")
        st.stop()

    mapping, missing = map_columns(df)
    if missing:
        st.error(f"Missing required column(s): **{', '.join(missing)}**. Found: {', '.join(df.columns)}")
        st.stop()

    st.subheader(f"Input preview ({len(df):,} leads)")
    st.dataframe(df[[mapping[c] for c in REQUIRED]].head(10), use_container_width=True)

    if st.button("Generate permutations", type="primary"):
        with st.spinner("Generating..."):
            out, skipped = build_output(
                df, mapping,
                mode="Validation (stricter, fewer patterns)" if not comprehensive else "Comprehensive",
                token_min_len=int(token_min_len),
                row_budget=int(row_budget),
                token_mode=token_mode,
                drop_existing=drop_existing,
            )
        st.session_state["result"] = (out, skipped, len(df))

if "result" in st.session_state:
    out, skipped, n_leads = st.session_state["result"]

    c1, c2, c3 = st.columns(3)
    c1.metric("Leads processed", f"{n_leads - len(skipped):,}")
    c2.metric("New emails generated", f"{len(out):,}")
    c3.metric("Leads skipped", f"{len(skipped):,}")

    st.subheader("Result preview")
    st.dataframe(out.head(1000), use_container_width=True)
    if len(out) > 1000:
        st.caption(f"Showing first 1,000 of {len(out):,} rows. The download has everything.")

    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇️ Download CSV",
        out.to_csv(index=False).encode("utf-8-sig"),
        file_name="email_permutations.csv",
        mime="text/csv",
    )
    if len(out) <= 1_000_000:
        buf = io.BytesIO()
        out.to_excel(buf, index=False, engine="openpyxl")
        d2.download_button(
            "⬇️ Download Excel",
            buf.getvalue(),
            file_name="email_permutations.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if len(skipped):
        with st.expander(f"{len(skipped)} lead(s) skipped (missing name/domain, or no valid name left after title cleanup)"):
            st.dataframe(skipped, use_container_width=True)
