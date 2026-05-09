#pragma once

#include "xlsx/xlsx_writer.hpp"
#include "xlsx/zip_file.hpp"
#include "xlsx/xml_util.hpp"
#include "xlsx/parsers/relationship_parser.hpp"
#include "xlsx/parsers/styles_append_parser.hpp"
#include "xlsx/parsers/workbook_append_parser.hpp"

#include "duckdb/common/file_system.hpp"
#include "duckdb/common/string_util.hpp"
#include "duckdb/common/exception.hpp"
#include "duckdb/main/client_context.hpp"

#include <atomic>
#include <chrono>
#include <unordered_set>

namespace duckdb {

class XLSXAppender {
public:
	// source_path: existing xlsx to read from (must exist)
	// dest_path:   where the new xlsx is written; caller (e.g. DuckDB COPY framework) is
	//              responsible for renaming dest_path onto the user-visible target. If the caller
	//              passes the same path for both, source is read fully before dest is opened, so
	//              same-path appends would lose data — don't do that without a tmp dance outside.
	XLSXAppender(ClientContext &context_p, const string &source_path_p, const string &dest_path_p,
	             const string &sheet_name_p, bool replace_p, idx_t sheet_row_limit_p);
	~XLSXAppender() = default;

	XLSXAppender(const XLSXAppender &) = delete;
	XLSXAppender &operator=(const XLSXAppender &) = delete;

	void BeginSheet(const vector<string> &sql_column_names, const vector<LogicalType> &sql_column_types);

	// Forwarded to the inner XLXSWriter
	void BeginRow() {
		writer->BeginRow();
	}
	void EndRow() {
		writer->EndRow();
	}
	void WriteNumberCell(const string_t &value) {
		writer->WriteNumberCell(value);
	}
	void WriteInlineStringCell(const string_t &value) {
		writer->WriteInlineStringCell(value);
	}
	void WriteBooleanCell(const string_t &value) {
		writer->WriteBooleanCell(value);
	}
	void WriteDateCell(const string_t &value) {
		writer->WriteDateCell(value);
	}
	void WriteTimeCell(const string_t &value) {
		writer->WriteTimeCell(value);
	}
	void WriteTimestampCell(const string_t &value) {
		writer->WriteTimestampCell(value);
	}
	void WriteTimestampCellNoMilliseconds(const string_t &value) {
		writer->WriteTimestampCellNoMilliseconds(value);
	}
	void WriteEmptyCell() {
		writer->WriteEmptyCell();
	}

	// End the sheet, write modified metadata, finalize the zip, atomic-rename onto the target.
	void Finish();

private:
	struct ResolvedStyles {
		idx_t date;
		idx_t ts_no_ms;
		idx_t time_;
		idx_t ts_with_ms;
		idx_t boolean;
		// numFmts and xf rows that need to be spliced into styles.xml. Empty == no patch needed.
		vector<string> num_fmt_inserts;
		vector<string> xf_inserts;
		// New count attribute values (only used when inserts are non-empty)
		idx_t new_num_fmts_count = 0;
		idx_t new_cell_xfs_count = 0;
		bool num_fmts_block_existed = false;
	};

	static string ReadEntryAsString(ZipFileReader &reader, const string &entry_name);
	static vector<char> ReadEntryAsBytes(ZipFileReader &reader, const string &entry_name);
	static idx_t ParseRidNumber(const string &rid);
	static string ToLowerAscii(const string &input);

	ResolvedStyles ResolveStyles(const string &raw_styles_xml, const StylesAppendParser &parser);
	string PatchStylesXml(const string &original, const ResolvedStyles &resolved);
	string BuildContentTypesXml(const string &original);
	string BuildWorkbookXml(const string &original);
	string BuildWorkbookRelsXml(const string &original);

	void EmitMetadata();

	ClientContext &context;
	string source_path;
	string dest_path;
	string sheet_name;
	bool replace;
	idx_t sheet_row_limit;

	// Discovery results
	vector<XLSXSheetEntry> source_sheets;
	vector<XLSXRelation> source_wb_rels;
	string source_content_types;
	string source_workbook_xml;
	string source_workbook_rels_xml;
	string source_styles_xml;
	bool source_has_styles = false;

	// Allocation for the new (or replacement) sheet
	string new_sheet_filename; // e.g. "sheetN.xml" (no path prefix)
	idx_t new_sheet_id = 0;
	idx_t new_rid_num = 0;
	bool replaced_existing = false;

	ResolvedStyles styles;
	bool patch_styles = false;

	// Set of source entries that get rewritten and must NOT be copied verbatim.
	std::unordered_set<string> rewrite_set;

	// Verbatim entries buffered from the source archive. Buffered up-front so we can
	// close the source file before opening the destination — Windows refuses to rename
	// over a file that still has any open handle on it (ERROR_SHARING_VIOLATION).
	struct BufferedEntry {
		string name;
		vector<char> bytes;
	};
	vector<BufferedEntry> verbatim_entries;

	unique_ptr<XLXSWriter> writer;
};

//===-- Implementation --------------------------------------------------------------------===//

inline vector<char> XLSXAppender::ReadEntryAsBytes(ZipFileReader &reader, const string &entry_name) {
	if (!reader.TryOpenEntry(entry_name)) {
		throw IOException("Required entry '%s' not found in xlsx file", entry_name);
	}
	const auto entry_len = reader.GetEntryLen();
	vector<char> result;
	result.resize(entry_len);
	idx_t pos = 0;
	constexpr idx_t BUFFER_SIZE = 4096;
	char buffer[BUFFER_SIZE];
	while (!reader.IsDone()) {
		const auto read_size = reader.Read(buffer, BUFFER_SIZE);
		if (read_size == 0) {
			break;
		}
		if (pos + read_size > result.size()) {
			result.resize(pos + read_size);
		}
		std::memcpy(result.data() + pos, buffer, read_size);
		pos += read_size;
	}
	if (pos < result.size()) {
		result.resize(pos);
	}
	reader.CloseEntry();
	return result;
}

inline string XLSXAppender::ReadEntryAsString(ZipFileReader &reader, const string &entry_name) {
	if (!reader.TryOpenEntry(entry_name)) {
		throw IOException("Required entry '%s' not found in xlsx file", entry_name);
	}
	const auto entry_len = reader.GetEntryLen();
	string result;
	result.resize(entry_len);
	idx_t pos = 0;
	constexpr idx_t BUFFER_SIZE = 4096;
	char buffer[BUFFER_SIZE];
	while (!reader.IsDone()) {
		const auto read_size = reader.Read(buffer, BUFFER_SIZE);
		if (read_size == 0) {
			break;
		}
		if (pos + read_size > entry_len) {
			result.resize(pos + read_size);
		}
		std::memcpy(&result[pos], buffer, read_size);
		pos += read_size;
	}
	if (pos < result.size()) {
		result.resize(pos);
	}
	reader.CloseEntry();
	return result;
}

inline idx_t XLSXAppender::ParseRidNumber(const string &rid) {
	// Expected pattern: "rId<digits>". Returns 0 for any non-conforming string.
	if (rid.size() < 4 || rid[0] != 'r' || rid[1] != 'I' || rid[2] != 'd') {
		return 0;
	}
	for (idx_t i = 3; i < rid.size(); i++) {
		if (rid[i] < '0' || rid[i] > '9') {
			return 0;
		}
	}
	return static_cast<idx_t>(std::strtoul(rid.c_str() + 3, nullptr, 10));
}

inline string XLSXAppender::ToLowerAscii(const string &input) {
	string result = input;
	for (auto &c : result) {
		if (c >= 'A' && c <= 'Z') {
			c = static_cast<char>(c - 'A' + 'a');
		}
	}
	return result;
}

inline XLSXAppender::XLSXAppender(ClientContext &context_p, const string &source_path_p, const string &dest_path_p,
                                  const string &sheet_name_p, bool replace_p, idx_t sheet_row_limit_p)
    : context(context_p), source_path(source_path_p), dest_path(dest_path_p), sheet_name(sheet_name_p),
      replace(replace_p), sheet_row_limit(sheet_row_limit_p) {

	// We deliberately scope the source ZipFileReader so its file handle is closed BEFORE the
	// destination writer is opened. On Windows the COPY framework's final rename fails if any
	// handle on the user-visible target file is still open (ERROR_SHARING_VIOLATION). All
	// data we'll need to copy verbatim is buffered into `verbatim_entries` inside this scope.
	{
	ZipFileReader source(context, source_path);

	// 2. Read raw bytes of the metadata files we'll rewrite. Every well-formed xlsx must have these.
	source_content_types = ReadEntryAsString(source, "[Content_Types].xml");
	source_workbook_xml = ReadEntryAsString(source, "xl/workbook.xml");
	source_workbook_rels_xml = ReadEntryAsString(source, "xl/_rels/workbook.xml.rels");

	// styles.xml is technically optional but ubiquitous in practice. We require it for the append
	// path because synthesizing one would also require splicing a Content_Types override for it.
	if (!source.TryOpenEntry("xl/styles.xml")) {
		throw IOException("Cannot append to '%s': xl/styles.xml is missing — only xlsx files with a styles.xml are supported",
		                  source_path);
	}
	source.CloseEntry();
	source_styles_xml = ReadEntryAsString(source, "xl/styles.xml");
	source_has_styles = true;

	// 3. Parse the structural data we need
	if (!source.TryOpenEntry("xl/workbook.xml")) {
		throw IOException("xl/workbook.xml missing in source");
	}
	source_sheets = WorkBookAppendParser::GetSheets(source);
	source.CloseEntry();

	if (!source.TryOpenEntry("xl/_rels/workbook.xml.rels")) {
		throw IOException("xl/_rels/workbook.xml.rels missing in source");
	}
	source_wb_rels = RelParser::ParseRelations(source);
	source.CloseEntry();

	// 4. Validate sheet name and decide what to write
	const auto target_lower = ToLowerAscii(sheet_name);
	idx_t collision_index = source_sheets.size();
	for (idx_t i = 0; i < source_sheets.size(); i++) {
		if (ToLowerAscii(source_sheets[i].name) == target_lower) {
			collision_index = i;
			break;
		}
	}

	if (replace) {
		if (collision_index == source_sheets.size()) {
			throw InvalidInputException("Sheet '%s' not found in '%s' (REPLACE)", sheet_name, source_path);
		}
		replaced_existing = true;
		// REPLACE: keep the existing sheet's rId, sheetId, and filename
		const auto &existing = source_sheets[collision_index];
		new_rid_num = ParseRidNumber(existing.rid);
		new_sheet_id = existing.sheet_id.empty() ? (collision_index + 1)
		                                         : static_cast<idx_t>(strtoul(existing.sheet_id.c_str(), nullptr, 10));

		// Resolve the worksheet target path via workbook.xml.rels
		string target_rel;
		for (auto &rel : source_wb_rels) {
			if (rel.id == existing.rid) {
				target_rel = rel.target;
				break;
			}
		}
		if (target_rel.empty()) {
			throw IOException("Could not resolve relationship '%s' for sheet '%s'", existing.rid, sheet_name);
		}
		// Normalize to a path inside the zip (relationships in xl/_rels are relative to xl/)
		string full_target;
		if (StringUtil::StartsWith(target_rel, "/")) {
			full_target = target_rel.substr(1);
		} else {
			full_target = "xl/" + target_rel;
		}
		// new_sheet_filename is just the leaf name; BeginSheet prepends "xl/worksheets/"
		auto last_slash = full_target.find_last_of('/');
		new_sheet_filename = (last_slash == string::npos) ? full_target : full_target.substr(last_slash + 1);

		// Verify this lives under xl/worksheets/ — if it doesn't, the path scheme differs and we'd
		// emit the new sheet at a different path than the old. Reject rather than corrupt the file.
		if (full_target != "xl/worksheets/" + new_sheet_filename) {
			throw IOException("REPLACE not supported: sheet '%s' is stored at '%s', expected xl/worksheets/<name>",
			                  sheet_name, full_target);
		}
		// REPLACE: skip the old sheet entry; everything else stays verbatim.
		rewrite_set.insert(full_target);
	} else {
		if (collision_index != source_sheets.size()) {
			throw InvalidInputException("Sheet '%s' already exists in '%s' (use REPLACE to overwrite)", sheet_name,
			                            source_path);
		}
		// APPEND: allocate fresh ids and a non-colliding filename
		idx_t max_sheet_id = 0;
		for (auto &sheet : source_sheets) {
			if (!sheet.sheet_id.empty()) {
				max_sheet_id = MaxValue<idx_t>(max_sheet_id,
				                               static_cast<idx_t>(strtoul(sheet.sheet_id.c_str(), nullptr, 10)));
			}
		}
		new_sheet_id = max_sheet_id + 1;

		idx_t max_rid = 0;
		for (auto &rel : source_wb_rels) {
			max_rid = MaxValue<idx_t>(max_rid, ParseRidNumber(rel.id));
		}
		new_rid_num = max_rid + 1;

		// Find a sheetN.xml that doesn't collide with anything in the source archive.
		const auto entry_names = source.ListEntries();
		std::unordered_set<string> entry_set(entry_names.begin(), entry_names.end());
		for (idx_t i = 1;; i++) {
			const auto candidate = "sheet" + std::to_string(i) + ".xml";
			const auto candidate_path = "xl/worksheets/" + candidate;
			if (entry_set.find(candidate_path) == entry_set.end()) {
				new_sheet_filename = candidate;
				break;
			}
		}
		// APPEND rewrites the workbook metadata too
		rewrite_set.insert("[Content_Types].xml");
		rewrite_set.insert("xl/workbook.xml");
		rewrite_set.insert("xl/_rels/workbook.xml.rels");
	}

	// 5. Resolve styles — find existing matching numFmts/cellXfs or queue up patches.
	if (source_has_styles) {
		StylesAppendParser styles_parser;
		if (source.TryOpenEntry("xl/styles.xml")) {
			styles_parser.ParseAll(source);
			source.CloseEntry();
		}
		styles = ResolveStyles(source_styles_xml, styles_parser);
	} else {
		// No styles.xml in source — synthesize one with our defaults (1..5).
		styles.date = 1;
		styles.ts_no_ms = 2;
		styles.time_ = 3;
		styles.ts_with_ms = 4;
		styles.boolean = 5;
		// Will need to write a fresh styles.xml; reuse XLXSWriter's WriteStyles equivalent below.
		// For simplicity we just ship the same default styles.xml the fresh-write path uses.
	}
	patch_styles = !styles.num_fmt_inserts.empty() || !styles.xf_inserts.empty() || !source_has_styles;
	if (patch_styles) {
		rewrite_set.insert("xl/styles.xml");
	}

	// 6. Buffer every entry that will be copied verbatim. Reading is fast and the data is
	// small (xlsx files are bounded; the cost is dwarfed by the cell-write loop).
	const auto entry_names = source.ListEntries();
	verbatim_entries.reserve(entry_names.size());
	for (auto &entry : entry_names) {
		if (rewrite_set.find(entry) != rewrite_set.end()) {
			continue;
		}
		// Skip directory entries (filename ends in '/'); the writer never needs them and
		// they confuse the BeginFile/Write/EndFile API.
		if (!entry.empty() && entry.back() == '/') {
			continue;
		}
		BufferedEntry buf;
		buf.name = entry;
		buf.bytes = ReadEntryAsBytes(source, entry);
		verbatim_entries.push_back(std::move(buf));
	}
	} // source destroyed here — file handle on source_path is released

	// 7. Now open the writer on the destination path and emit the buffered entries.
	// At this point no handle on source_path remains open, so the COPY framework's final
	// rename of dest_path onto source_path will succeed even on Windows.
	writer = make_uniq<XLXSWriter>(context, dest_path, sheet_row_limit);
	writer->SetStyleIndices(styles.date, styles.ts_no_ms, styles.time_, styles.ts_with_ms, styles.boolean);

	for (auto &entry : verbatim_entries) {
		auto &out = writer->GetStream();
		out.BeginFile(entry.name);
		if (!entry.bytes.empty()) {
			out.Write(entry.bytes.data(), entry.bytes.size());
		}
		out.EndFile();
	}
	verbatim_entries.clear();  // free memory; we no longer need the buffers
}

inline void XLSXAppender::BeginSheet(const vector<string> &sql_column_names,
                                     const vector<LogicalType> &sql_column_types) {
	writer->BeginSheet(sheet_name, new_sheet_filename, sql_column_names, sql_column_types);
}

inline XLSXAppender::ResolvedStyles XLSXAppender::ResolveStyles(const string &raw_styles_xml,
                                                                const StylesAppendParser &parser) {
	(void)raw_styles_xml; // unused; kept for future use
	ResolvedStyles result;

	// The 5 number formats this writer needs. `format_code` is the *unescaped* form (what expat
	// hands us when parsing styles.xml — entities like &quot; are already decoded). Comparison
	// against parsed entries uses this form. When we splice a fresh <numFmt> back into styles.xml
	// we re-escape via XML entities.
	struct Needed {
		const char *format_code;       // unescaped, for parser-side comparison
		const char *format_code_xml;   // XML-escaped, for writing back into styles.xml
		idx_t default_id;              // unused but kept for documentation
		idx_t *out_index;
	};
	Needed needed[] = {
	    {"dd/mm/yy",                 "dd/mm/yy",                                                       165, &result.date},
	    {"dd/mm/yyyy\\ hh:mm:ss",    "dd/mm/yyyy\\ hh:mm:ss",                                          166, &result.ts_no_ms},
	    {"hh:mm:ss",                 "hh:mm:ss",                                                       167, &result.time_},
	    {"dd/mm/yyyy\\ hh:mm:ss.000","dd/mm/yyyy\\ hh:mm:ss.000",                                      168, &result.ts_with_ms},
	    {"\"TRUE\";\"TRUE\";\"FALSE\"", "&quot;TRUE&quot;;&quot;TRUE&quot;;&quot;FALSE&quot;",         169, &result.boolean},
	};

	const auto initial_num_fmts = parser.num_fmts.size();
	const auto initial_cell_xfs = parser.cell_xfs.size();
	result.num_fmts_block_existed = initial_num_fmts > 0;

	// Track running additions
	vector<XLSXNumFmtEntry> effective_num_fmts = parser.num_fmts;
	vector<XLSXCellXfEntry> effective_cell_xfs = parser.cell_xfs;

	idx_t next_custom_id = 165;
	for (auto &entry : parser.num_fmts) {
		if (entry.num_fmt_id >= next_custom_id) {
			next_custom_id = entry.num_fmt_id + 1;
		}
	}

	for (auto &need : needed) {
		// Resolve numFmtId for this format code (find existing or queue insertion)
		idx_t resolved_num_fmt_id = 0;
		bool found_num_fmt = false;
		for (auto &nfmt : effective_num_fmts) {
			if (nfmt.format_code == need.format_code) {
				resolved_num_fmt_id = nfmt.num_fmt_id;
				found_num_fmt = true;
				break;
			}
		}
		if (!found_num_fmt) {
			resolved_num_fmt_id = next_custom_id++;
			XLSXNumFmtEntry new_entry;
			new_entry.num_fmt_id = resolved_num_fmt_id;
			new_entry.format_code = need.format_code;
			effective_num_fmts.push_back(new_entry);
			const auto element = StringUtil::Format(R"(<numFmt formatCode="%s" numFmtId="%d"/>)",
			                                        need.format_code_xml, resolved_num_fmt_id);
			result.num_fmt_inserts.push_back(element);
		}

		// Resolve cellXfs index for this numFmtId (find existing or queue insertion)
		idx_t resolved_xf_idx = 0;
		bool found_xf = false;
		for (idx_t i = 0; i < effective_cell_xfs.size(); i++) {
			if (effective_cell_xfs[i].num_fmt_id == resolved_num_fmt_id) {
				resolved_xf_idx = i;
				found_xf = true;
				break;
			}
		}
		if (!found_xf) {
			resolved_xf_idx = effective_cell_xfs.size();
			XLSXCellXfEntry new_xf;
			new_xf.num_fmt_id = resolved_num_fmt_id;
			effective_cell_xfs.push_back(new_xf);
			const auto element = StringUtil::Format(R"(<xf numFmtId="%d" xfId="0"/>)", resolved_num_fmt_id);
			result.xf_inserts.push_back(element);
		}
		*need.out_index = resolved_xf_idx;
	}

	result.new_num_fmts_count = effective_num_fmts.size();
	result.new_cell_xfs_count = effective_cell_xfs.size();
	(void)initial_num_fmts;
	(void)initial_cell_xfs;
	return result;
}

inline string XLSXAppender::PatchStylesXml(const string &original, const ResolvedStyles &resolved) {
	string patched = original;

	// Splice xf rows before </cellXfs> and update count attribute
	if (!resolved.xf_inserts.empty()) {
		const auto close_pos = patched.rfind("</cellXfs>");
		if (close_pos == string::npos) {
			throw IOException("Cannot patch styles.xml: missing </cellXfs>");
		}
		string inserts;
		for (auto &xf : resolved.xf_inserts) {
			inserts += xf;
		}
		patched.insert(close_pos, inserts);

		// Update the count attribute on the <cellXfs ...> opening tag
		const auto open_pos = patched.find("<cellXfs");
		if (open_pos != string::npos) {
			const auto tag_end = patched.find('>', open_pos);
			if (tag_end != string::npos) {
				auto count_pos = patched.find("count=\"", open_pos);
				if (count_pos != string::npos && count_pos < tag_end) {
					const auto count_val_start = count_pos + strlen("count=\"");
					const auto count_val_end = patched.find('"', count_val_start);
					if (count_val_end != string::npos) {
						patched.replace(count_val_start, count_val_end - count_val_start,
						                std::to_string(resolved.new_cell_xfs_count));
					}
				}
			}
		}
	}

	// Splice numFmt rows before </numFmts> and update count attribute (or create the block)
	if (!resolved.num_fmt_inserts.empty()) {
		string inserts;
		for (auto &nfmt : resolved.num_fmt_inserts) {
			inserts += nfmt;
		}
		const auto close_pos = patched.rfind("</numFmts>");
		if (close_pos != string::npos) {
			patched.insert(close_pos, inserts);
			// Update count
			const auto open_pos = patched.find("<numFmts");
			if (open_pos != string::npos) {
				const auto tag_end = patched.find('>', open_pos);
				if (tag_end != string::npos) {
					auto count_pos = patched.find("count=\"", open_pos);
					if (count_pos != string::npos && count_pos < tag_end) {
						const auto count_val_start = count_pos + strlen("count=\"");
						const auto count_val_end = patched.find('"', count_val_start);
						if (count_val_end != string::npos) {
							patched.replace(count_val_start, count_val_end - count_val_start,
							                std::to_string(resolved.new_num_fmts_count));
						}
					}
				}
			}
		} else {
			// No <numFmts> block exists — create one right after the styleSheet opening tag
			const auto sheet_open = patched.find("<styleSheet");
			if (sheet_open == string::npos) {
				throw IOException("Cannot patch styles.xml: missing <styleSheet>");
			}
			const auto sheet_close = patched.find('>', sheet_open);
			if (sheet_close == string::npos) {
				throw IOException("Cannot patch styles.xml: malformed <styleSheet>");
			}
			const auto block = StringUtil::Format(R"(<numFmts count="%d">%s</numFmts>)",
			                                      resolved.num_fmt_inserts.size(), inserts);
			patched.insert(sheet_close + 1, block);
		}
	}

	return patched;
}

inline string XLSXAppender::BuildContentTypesXml(const string &original) {
	// Splice a single <Override> for the new sheet just before </Types>
	const auto close_pos = original.rfind("</Types>");
	if (close_pos == string::npos) {
		throw IOException("Cannot patch [Content_Types].xml: missing </Types>");
	}
	const auto override_xml = StringUtil::Format(
	    R"(<Override PartName="/xl/worksheets/%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>)",
	    new_sheet_filename);
	string result = original;
	result.insert(close_pos, override_xml);
	return result;
}

inline string XLSXAppender::BuildWorkbookXml(const string &original) {
	const auto close_pos = original.rfind("</sheets>");
	if (close_pos == string::npos) {
		throw IOException("Cannot patch xl/workbook.xml: missing </sheets>");
	}
	const auto sheet_xml = StringUtil::Format(R"(<sheet name="%s" state="visible" sheetId="%d" r:id="rId%d"/>)",
	                                          EscapeXMLString(sheet_name), new_sheet_id, new_rid_num);
	string result = original;
	result.insert(close_pos, sheet_xml);
	return result;
}

inline string XLSXAppender::BuildWorkbookRelsXml(const string &original) {
	const auto close_pos = original.rfind("</Relationships>");
	if (close_pos == string::npos) {
		throw IOException("Cannot patch xl/_rels/workbook.xml.rels: missing </Relationships>");
	}
	const auto rel_xml = StringUtil::Format(
	    R"(<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/%s"/>)",
	    new_rid_num, new_sheet_filename);
	string result = original;
	result.insert(close_pos, rel_xml);
	return result;
}

inline void XLSXAppender::EmitMetadata() {
	auto &out = writer->GetStream();

	if (!replaced_existing) {
		// APPEND: rewrite Content_Types, workbook.xml, workbook.xml.rels
		const auto content_types = BuildContentTypesXml(source_content_types);
		out.BeginFile("[Content_Types].xml");
		out.Write(content_types);
		out.EndFile();

		const auto workbook = BuildWorkbookXml(source_workbook_xml);
		out.BeginFile("xl/workbook.xml");
		out.Write(workbook);
		out.EndFile();

		const auto workbook_rels = BuildWorkbookRelsXml(source_workbook_rels_xml);
		out.BeginFile("xl/_rels/workbook.xml.rels");
		out.Write(workbook_rels);
		out.EndFile();
	}
	// REPLACE doesn't touch [Content_Types].xml / workbook.xml / workbook.xml.rels;
	// the verbatim copy already preserved them.

	if (patch_styles) {
		string styles_xml;
		if (source_has_styles) {
			styles_xml = PatchStylesXml(source_styles_xml, styles);
		} else {
			// Fresh styles.xml using our defaults — inline the same content as XLXSWriter::WriteStyles.
			styles_xml =
			    R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
			    R"(<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">)"
			    R"(<numFmts count="6">)"
			    R"(<numFmt formatCode="General" numFmtId="164"/>)"
			    R"(<numFmt formatCode="dd/mm/yy" numFmtId="165"/>)"
			    R"(<numFmt formatCode="dd/mm/yyyy\ hh:mm:ss" numFmtId="166"/>)"
			    R"(<numFmt formatCode="hh:mm:ss" numFmtId="167"/>)"
			    R"(<numFmt formatCode="dd/mm/yyyy\ hh:mm:ss.000" numFmtId="168"/>)"
			    R"(<numFmt formatCode="&quot;TRUE&quot;;&quot;TRUE&quot;;&quot;FALSE&quot;" numFmtId="169"/>)"
			    R"(</numFmts>)"
			    R"(<fonts count="1"><font><name val="Arial"/><family val="2"/><sz val="12"/></font></fonts>)"
			    R"(<fills count="1"><fill><patternFill patternType="none"/></fill></fills>)"
			    R"(<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>)"
			    R"(<cellStyleXfs count="1"><xf numFmtId="164"></xf></cellStyleXfs>)"
			    R"(<cellXfs count="6">)"
			    R"(<xf numFmtId="164" xfId="0"/>)"
			    R"(<xf numFmtId="165" xfId="0"/>)"
			    R"(<xf numFmtId="166" xfId="0"/>)"
			    R"(<xf numFmtId="167" xfId="0"/>)"
			    R"(<xf numFmtId="168" xfId="0"/>)"
			    R"(<xf numFmtId="169" xfId="0"/>)"
			    R"(</cellXfs>)"
			    R"(<cellStyles count="1"><cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/></cellStyles>)"
			    R"(</styleSheet>)";
		}
		out.BeginFile("xl/styles.xml");
		out.Write(styles_xml);
		out.EndFile();
	}
}

inline void XLSXAppender::Finish() {
	writer->EndSheet();
	EmitMetadata();
	writer->GetStream().Finalize();
	// dest_path is now a complete xlsx; the caller (DuckDB COPY framework) will rename it onto the
	// user-visible target path.
}

} // namespace duckdb
