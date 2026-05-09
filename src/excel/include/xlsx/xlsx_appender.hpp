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
#include <cstring>
#include <unordered_set>

namespace duckdb {

// Append a new sheet (or replace one) into an existing xlsx by rebuilding the zip with
// fast raw-copy of unchanged entries (no decompression / recompression) plus the new
// sheet and updated metadata files.
//
// Flow:
//   1. Open the source xlsx and read just the metadata files we need (workbook.xml,
//      workbook.xml.rels, [Content_Types].xml, styles.xml) — typically a few KB.
//   2. Decide which source entries we'll replace (the metadata files + the target sheet
//      for REPLACE mode). Build a "skip set".
//   3. Open the destination zip writer in CREATE mode (writes a fresh zip).
//   4. Walk source entries; for each one not in the skip set (and the canonical / last
//      occurrence if there are duplicates from a previously appended file), call
//      `mz_zip_writer_copy_from_reader` — pure raw-byte stream copy, no decompression.
//   5. Write the new sheet's XML and the updated metadata files into the destination.
//
// Per-append cost: O(compressed_source_bytes) for the raw stream-copy + O(new_sheet_size)
// for the new sheet. Independent of how many cells the existing sheets have, because no
// existing sheet is ever decompressed.
class XLSXAppender {
public:
	// source_path: existing xlsx to read entries from (must exist). All entries except
	//              the ones we explicitly replace are stream-copied verbatim from here.
	// dest_path:   where to write the resulting xlsx (a fresh zip). Must NOT equal
	//              source_path — we open dest in CREATE mode (truncate). DuckDB's COPY
	//              framework satisfies this by handing us its own temp file path.
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

	// Close the new sheet, append updated metadata, finalize the zip.
	void Finish();

private:
	struct ResolvedStyles {
		idx_t date = 1;
		idx_t ts_no_ms = 2;
		idx_t time_ = 3;
		idx_t ts_with_ms = 4;
		idx_t boolean = 5;
		// numFmts and xf rows that need to be spliced into a replacement styles.xml.
		// Empty == no patch needed (existing styles.xml already covers everything we use).
		vector<string> num_fmt_inserts;
		vector<string> xf_inserts;
		idx_t new_num_fmts_count = 0;
		idx_t new_cell_xfs_count = 0;
	};

	static string ReadEntryAsString(ZipFileReader &reader, const string &entry_name);
	static idx_t ParseRidNumber(const string &rid);
	static string ToLowerAscii(const string &input);

	ResolvedStyles ResolveStyles(const StylesAppendParser &parser);
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

	// Discovery results — raw bytes of the metadata files, kept for the splice.
	vector<XLSXSheetEntry> source_sheets;
	vector<XLSXRelation> source_wb_rels;
	string source_content_types;
	string source_workbook_xml;
	string source_workbook_rels_xml;
	string source_styles_xml;

	// Allocation for the new (or replacement) sheet
	string new_sheet_filename; // e.g. "sheetN.xml" (no path prefix; BeginSheet prepends "xl/worksheets/")
	idx_t new_sheet_id = 0;
	idx_t new_rid_num = 0;
	bool replaced_existing = false;

	ResolvedStyles styles;
	bool styles_need_patch = false;

	// Source entries we will NOT verbatim-copy because we'll write a fresh version.
	std::unordered_set<string> rewrite_set;

	// When source_path == dest_path (e.g. DuckDB's COPY framework set USE_TMP_FILE=false
	// because its FileExists check on a path with `~` returned false), we can't write
	// directly to dest_path without losing the source we still need to read. In that case
	// we write to our own sibling temp and rename onto dest_path in Finish.
	string actual_write_path;
	bool needs_self_rename = false;

	unique_ptr<XLXSWriter> writer;
};

//===-- Implementation --------------------------------------------------------------------===//

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
		if (pos + read_size > result.size()) {
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

	if (source_path == dest_path) {
		// Pick a unique sibling tmp path; rename ourselves in Finish.
		static std::atomic<uint64_t> tmp_counter {0};
		const auto now_ns =
		    static_cast<uint64_t>(std::chrono::high_resolution_clock::now().time_since_epoch().count());
		const auto seq = tmp_counter.fetch_add(1);
		actual_write_path = source_path + ".xlsxapp.tmp." + std::to_string(now_ns) + "." + std::to_string(seq);
		needs_self_rename = true;
	} else {
		actual_write_path = dest_path;
	}

	// Phase 1: read structural metadata from source. Kept in a tight scope so the source
	// file handle is released before we proceed to file-copy or append-mode open.
	{
		ZipFileReader source(context, source_path);

		source_content_types = ReadEntryAsString(source, "[Content_Types].xml");
		source_workbook_xml = ReadEntryAsString(source, "xl/workbook.xml");
		source_workbook_rels_xml = ReadEntryAsString(source, "xl/_rels/workbook.xml.rels");

		if (!source.TryOpenEntry("xl/styles.xml")) {
			throw IOException("Cannot append to '%s': xl/styles.xml is missing", source_path);
		}
		source.CloseEntry();
		source_styles_xml = ReadEntryAsString(source, "xl/styles.xml");

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

		StylesAppendParser styles_parser;
		if (source.TryOpenEntry("xl/styles.xml")) {
			styles_parser.ParseAll(source);
			source.CloseEntry();
		}
		styles = ResolveStyles(styles_parser);
	}
	// source ZipFileReader destroyed here — file handle released.

	styles_need_patch = !styles.num_fmt_inserts.empty() || !styles.xf_inserts.empty();

	// Phase 2: validate sheet name and allocate the new sheet's identifiers / filename.
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
		// Reuse the existing sheet's rId, sheetId, and filename. The new worksheet entry
		// gets the same name as the old one — last-wins semantics make our content active
		// and external references (charts, defined names, pivots) keep resolving.
		const auto &existing = source_sheets[collision_index];
		new_rid_num = ParseRidNumber(existing.rid);
		new_sheet_id = existing.sheet_id.empty() ? (collision_index + 1)
		                                         : static_cast<idx_t>(strtoul(existing.sheet_id.c_str(), nullptr, 10));

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
		string full_target = StringUtil::StartsWith(target_rel, "/") ? target_rel.substr(1) : ("xl/" + target_rel);
		auto last_slash = full_target.find_last_of('/');
		new_sheet_filename = (last_slash == string::npos) ? full_target : full_target.substr(last_slash + 1);

		if (full_target != "xl/worksheets/" + new_sheet_filename) {
			throw IOException("REPLACE not supported: sheet '%s' is stored at '%s', expected xl/worksheets/<name>",
			                  sheet_name, full_target);
		}
		// REPLACE: skip the existing sheet's worksheet entry — we'll rewrite it.
		// All other entries (including workbook.xml / rels / content_types) are kept.
		rewrite_set.insert(full_target);
	} else {
		if (collision_index != source_sheets.size()) {
			throw InvalidInputException("Sheet '%s' already exists in '%s' (use REPLACE to overwrite)", sheet_name,
			                            source_path);
		}
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

		// Allocate a non-colliding sheet filename. We don't have the entry list here (we
		// closed the source reader to release the handle), but we can derive it from the
		// existing relationships: every worksheet relationship points at a `worksheets/sheetN.xml`.
		std::unordered_set<idx_t> used_indices;
		for (auto &rel : source_wb_rels) {
			if (!StringUtil::EndsWith(rel.type, "/worksheet")) {
				continue;
			}
			// Extract N from "worksheets/sheetN.xml" (or "/xl/worksheets/sheetN.xml").
			auto target = rel.target;
			auto last_slash = target.find_last_of('/');
			auto leaf = (last_slash == string::npos) ? target : target.substr(last_slash + 1);
			if (StringUtil::StartsWith(leaf, "sheet") && StringUtil::EndsWith(leaf, ".xml")) {
				auto digits = leaf.substr(5, leaf.size() - 5 - 4);
				bool all_digits = !digits.empty();
				for (auto c : digits) {
					if (c < '0' || c > '9') {
						all_digits = false;
						break;
					}
				}
				if (all_digits) {
					used_indices.insert(static_cast<idx_t>(std::strtoul(digits.c_str(), nullptr, 10)));
				}
			}
		}
		for (idx_t i = 1;; i++) {
			if (used_indices.find(i) == used_indices.end()) {
				new_sheet_filename = "sheet" + std::to_string(i) + ".xml";
				break;
			}
		}
		// APPEND rewrites the workbook metadata so the new sheet is referenced
		rewrite_set.insert("[Content_Types].xml");
		rewrite_set.insert("xl/workbook.xml");
		rewrite_set.insert("xl/_rels/workbook.xml.rels");
	}
	if (styles_need_patch) {
		rewrite_set.insert("xl/styles.xml");
	}

	// Phase 3: open dest as a fresh zip and stream-copy the entries we want to keep.
	// `mz_zip_writer_copy_from_reader` copies raw compressed bytes — no decompression
	// or recompression. Cost is O(compressed_source_bytes) of pure I/O.
	auto &fs = FileSystem::GetFileSystem(context);
	if (fs.FileExists(actual_write_path)) {
		fs.RemoveFile(actual_write_path);
	}
	writer = make_uniq<XLXSWriter>(context, actual_write_path, sheet_row_limit);
	writer->SetStyleIndices(styles.date, styles.ts_no_ms, styles.time_, styles.ts_with_ms, styles.boolean);

	{
		ZipFileReader source(context, source_path);

		// Build a name -> last-occurrence-index map so that if the source has duplicate-name
		// entries (e.g. files written by an older buggy version of the appender), we copy the
		// authoritative (last) one and discard earlier shadows.
		const auto names = source.ListEntries();
		std::unordered_map<string, idx_t> last_index_for;
		for (idx_t i = 0; i < names.size(); i++) {
			last_index_for[names[i]] = i;
		}

		idx_t walk_idx = 0;
		if (source.GotoFirstEntry()) {
			do {
				const auto &name = names[walk_idx];
				const bool is_canonical = last_index_for[name] == walk_idx;
				const bool is_skipped = rewrite_set.count(name) > 0;
				const bool is_dir = !name.empty() && name.back() == '/';
				if (is_canonical && !is_skipped && !is_dir) {
					writer->GetStream().CopyCurrentEntryFrom(source);
				}
				walk_idx++;
			} while (source.GotoNextEntry());
		}
	}
	// source ZipFileReader released here; writer holds dest open until Finish().
}

inline void XLSXAppender::BeginSheet(const vector<string> &sql_column_names,
                                     const vector<LogicalType> &sql_column_types) {
	writer->BeginSheet(sheet_name, new_sheet_filename, sql_column_names, sql_column_types);
}

inline XLSXAppender::ResolvedStyles XLSXAppender::ResolveStyles(const StylesAppendParser &parser) {
	ResolvedStyles result;

	// The 5 number formats this writer needs. format_code is the *unescaped* form (what
	// expat hands us when parsing styles.xml — entities like &quot; are already decoded).
	// format_code_xml is the XML-escaped form used when writing back into styles.xml.
	struct Needed {
		const char *format_code;
		const char *format_code_xml;
		idx_t *out_index;
	};
	Needed needed[] = {
	    {"dd/mm/yy",                     "dd/mm/yy",                                                       &result.date},
	    {"dd/mm/yyyy\\ hh:mm:ss",        "dd/mm/yyyy\\ hh:mm:ss",                                          &result.ts_no_ms},
	    {"hh:mm:ss",                     "hh:mm:ss",                                                       &result.time_},
	    {"dd/mm/yyyy\\ hh:mm:ss.000",    "dd/mm/yyyy\\ hh:mm:ss.000",                                      &result.ts_with_ms},
	    {"\"TRUE\";\"TRUE\";\"FALSE\"",  "&quot;TRUE&quot;;&quot;TRUE&quot;;&quot;FALSE&quot;",            &result.boolean},
	};

	vector<XLSXNumFmtEntry> effective_num_fmts = parser.num_fmts;
	vector<XLSXCellXfEntry> effective_cell_xfs = parser.cell_xfs;

	idx_t next_custom_id = 165;
	for (auto &entry : parser.num_fmts) {
		if (entry.num_fmt_id >= next_custom_id) {
			next_custom_id = entry.num_fmt_id + 1;
		}
	}

	for (auto &need : needed) {
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
			effective_num_fmts.push_back({resolved_num_fmt_id, need.format_code});
			result.num_fmt_inserts.push_back(StringUtil::Format(
			    R"(<numFmt formatCode="%s" numFmtId="%d"/>)", need.format_code_xml, resolved_num_fmt_id));
		}

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
			effective_cell_xfs.push_back({resolved_num_fmt_id});
			result.xf_inserts.push_back(
			    StringUtil::Format(R"(<xf numFmtId="%d" xfId="0"/>)", resolved_num_fmt_id));
		}
		*need.out_index = resolved_xf_idx;
	}

	result.new_num_fmts_count = effective_num_fmts.size();
	result.new_cell_xfs_count = effective_cell_xfs.size();
	return result;
}

inline string XLSXAppender::PatchStylesXml(const string &original, const ResolvedStyles &resolved) {
	string patched = original;

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

	if (!resolved.num_fmt_inserts.empty()) {
		string inserts;
		for (auto &nfmt : resolved.num_fmt_inserts) {
			inserts += nfmt;
		}
		const auto close_pos = patched.rfind("</numFmts>");
		if (close_pos != string::npos) {
			patched.insert(close_pos, inserts);
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
		// APPEND: write replacement workbook.xml, workbook.xml.rels, [Content_Types].xml.
		// These appear later in the central directory than the originals, so last-wins
		// readers (Excel, openpyxl, our own read_xlsx) pick up our updated versions.
		out.BeginFile("[Content_Types].xml");
		out.Write(BuildContentTypesXml(source_content_types));
		out.EndFile();

		out.BeginFile("xl/workbook.xml");
		out.Write(BuildWorkbookXml(source_workbook_xml));
		out.EndFile();

		out.BeginFile("xl/_rels/workbook.xml.rels");
		out.Write(BuildWorkbookRelsXml(source_workbook_rels_xml));
		out.EndFile();
	}
	// REPLACE keeps workbook.xml / workbook.xml.rels / [Content_Types].xml unchanged —
	// the new sheet entry shares the existing sheet's filename, so last-wins handles it.

	if (styles_need_patch) {
		out.BeginFile("xl/styles.xml");
		out.Write(PatchStylesXml(source_styles_xml, styles));
		out.EndFile();
	}
}

inline void XLSXAppender::Finish() {
	writer->EndSheet();
	EmitMetadata();
	writer->GetStream().Finalize();
	// `actual_write_path` is now a complete xlsx. If we wrote to a sibling tmp because
	// source_path == dest_path, rename our tmp onto dest_path now. Otherwise the COPY
	// framework will handle the rename (its own tmp -> user target).
	if (needs_self_rename) {
		auto &fs = FileSystem::GetFileSystem(context);
		try {
			if (fs.FileExists(dest_path)) {
				fs.RemoveFile(dest_path);
			}
			fs.MoveFile(actual_write_path, dest_path);
		} catch (...) {
			// Best-effort cleanup of the orphan tmp; let the original error propagate.
			try {
				if (fs.FileExists(actual_write_path)) {
					fs.RemoveFile(actual_write_path);
				}
			} catch (...) {
			}
			throw;
		}
	}
}

} // namespace duckdb
