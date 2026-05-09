#pragma once

#include "duckdb/common/typedefs.hpp"
#include "duckdb/common/string.hpp"
#include "duckdb/common/vector.hpp"

namespace duckdb {

class ClientContext;

class ZipFileReader;

class ZipFileWriter {
public:
	ZipFileWriter(ClientContext &context, const string &file_name);
	~ZipFileWriter();

	// Delete copy
	ZipFileWriter(const ZipFileWriter &) = delete;
	ZipFileWriter &operator=(const ZipFileWriter &) = delete;

	void AddDirectory(const string &dir_name);
	void BeginFile(const string &file_name);
	idx_t Write(const char *buffer, idx_t write_size);
	idx_t Write(const string &str);
	idx_t Write(const char *str);

	void EndFile();
	void Finalize();

	// Copy the source reader's CURRENTLY-POSITIONED entry verbatim into this writer.
	// Uses minizip's raw-copy primitive — no decompression or recompression. Caller must
	// position the reader (e.g. via GotoFirstEntry / GotoNextEntry) before calling.
	void CopyCurrentEntryFrom(ZipFileReader &source);

private:
	void *handle;
	void *stream;
	bool is_entry_open;
	vector<char> escaped_buffer;
};

class ZipFileReader {
public:
	ZipFileReader(ClientContext &context, const string &file_name);
	~ZipFileReader();

	// Delete copy
	ZipFileReader(const ZipFileReader &) = delete;
	ZipFileReader &operator=(const ZipFileReader &) = delete;

	bool TryOpenEntry(const string &file_name);
	void CloseEntry();
	idx_t Read(char *buffer, idx_t read_size);

	// Returns the current position in the current entry
	idx_t GetEntryPos() const;
	// Returns the uncompressed size of the current entry
	idx_t GetEntryLen() const;
	// Returns if the current entry is done
	bool IsDone() const;

	// Walk all entries in the archive and return their names. The reader must not have an entry currently open.
	vector<string> ListEntries();

	// Returns true if the entry exists and is a directory entry (filename ends with '/').
	bool EntryIsDirectory(const string &entry_name);

	// Position the cursor for sequential entry walking. Returns false when no more entries
	// (or on error). After GotoFirstEntry / GotoNextEntry returns true, the reader is positioned
	// on an entry and CurrentEntryName / CurrentEntryIsDirectory / ZipFileWriter::CopyCurrentEntryFrom
	// can be used.
	bool GotoFirstEntry();
	bool GotoNextEntry();
	string CurrentEntryName();
	bool CurrentEntryIsDirectory();

private:
	friend class ZipFileWriter;

	void *handle;
	void *stream;
	bool is_entry_open;

	idx_t entry_pos;
	idx_t entry_len;
};

} // namespace duckdb