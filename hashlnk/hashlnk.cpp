#define NOMINMAX
#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include <Shellapi.h>
#include <Shlwapi.h>
#include <Shlobj.h>
#include <Knownfolders.h>
#include <propkey.h>

void* __cdecl operator new(size_t n) {
	return HeapAlloc(GetProcessHeap(), 0, n);
}

void __cdecl operator delete(void* p) {
	if (p != nullptr) {
		HeapFree(GetProcessHeap(), 0, p);
	}
}

void* __cdecl operator new[](size_t n) {
	return operator new(n);
}

void __cdecl operator delete[](void* p) {
	operator delete(p);
}

template <typename T>
T* Allocate(size_t n) {
	return reinterpret_cast<T*>(HeapAlloc(GetProcessHeap(), 0, sizeof(T) * n));
}

template <typename T>
T* Reallocate(void* p, size_t n) {
	return reinterpret_cast<T*>(HeapReAlloc(GetProcessHeap(), 0, p, sizeof(T) * n));
}

void Free(void* p) {
	HeapFree(GetProcessHeap(), 0, p);
}

const PROPERTYKEY PKEY_WinX_Hash = { { 0xFB8D2D7B, 0x90D1, 0x4E34, { 0xBF, 0x60, 0x6E, 0xAC, 0x09, 0x92, 0x2B, 0xBF } }, 0x02 };

DWORD HResultToWin32Error(HRESULT hres) {
	if ((hres & 0xFFFF0000) == MAKE_HRESULT(SEVERITY_ERROR, FACILITY_WIN32, 0)) {
		return HRESULT_CODE(hres);
	}

	if (hres == S_OK) {
		return ERROR_SUCCESS;
	}

	return ERROR_CAN_NOT_COMPLETE;
}

template <typename T>
class LocalPtr {
public:
	LocalPtr() : ptr(nullptr) { }
	~LocalPtr() { if (ptr != nullptr) LocalFree(ptr); }
	LocalPtr(LocalPtr&& other) : ptr(other.ptr) { other.ptr = nullptr; }
	LocalPtr(const LocalPtr&) = delete;
	LocalPtr& operator=(const LocalPtr&) = delete;
	T* get() { return ptr; }
	T** operator&() { return &ptr; }
private:
	T* ptr;
};


template <typename T>
class CoTaskPtr {
public:
	CoTaskPtr() : ptr(nullptr) {}
	~CoTaskPtr() { if (ptr != nullptr) CoTaskMemFree(ptr); }
	CoTaskPtr(const CoTaskPtr&) = delete;
	CoTaskPtr& operator=(const CoTaskPtr&) = delete;
	CoTaskPtr(CoTaskPtr&& other) : ptr(other.ptr) { other.ptr = nullptr; }
	T* get() { return ptr; }
	T** operator&() { return &ptr; }
private:
	T* ptr;
};

template <typename T>
class ComPtr {
public:
	ComPtr() : ptr(nullptr) {}
	~ComPtr() { if (ptr != nullptr) 	ptr->Release();	}
	ComPtr(const ComPtr& other) : ptr(other.ptr) { ptr->AddRef(); }
	ComPtr(ComPtr&& other) : ptr(other.ptr) { other.ptr = nullptr; }

	ComPtr& operator=(const ComPtr& other) {
		if (&other != this) {
			ptr = other.ptr;
			ptr->AddRef();
		}
		return *this;
	}

	T* get() { return ptr; }
	T* operator->() { return ptr; }
	T** operator&() { return &ptr; }
private:
	T* ptr;
};

template <typename T>
size_t StringLength(const T* str) {
	size_t len = 0;
	for (; str[len] != 0; ++len) {}
	return len;
}

template <typename T>
wchar_t* CopyString(T* dest, const T* src) {
	auto ret = dest;
	while ((*dest++ = *src++)) {}
	return ret;
}

// Very basic implementation of a string
class wstring {
public:
	wstring() : cap(1), ptr(Allocate<wchar_t>(cap)) { ptr[0] = L'\0'; }
	wstring(const wstring& other) : cap(other.cap), ptr(Allocate<wchar_t>(cap)) {
		CopyString(ptr, other.ptr);
	}
	wstring(wstring&& other) : cap(other.cap), ptr(other.ptr) {
		other.ptr = nullptr;
	}
	wstring(const wchar_t* other) : cap(StringLength(other) + 1), ptr(Allocate<wchar_t>(cap)) {
		CopyString(ptr, other);
	}
	~wstring() { if (ptr != nullptr) Free(ptr); }
	wstring& append(const wstring& other) {
		return append(other.ptr, other.size());
	}
	wstring& append(const wchar_t* other, size_t otherLen) {
		ptr = Reallocate<wchar_t>(ptr, cap + otherLen);
		CopyString(ptr + size(), other);
		cap += otherLen;
		return *this;
	}
	wstring& append(const wchar_t* other) {
		return append(other, StringLength(other));
	}
	size_t size() const { return cap - 1; }
	const wchar_t* data() const { return ptr; }
	wchar_t* data() { return ptr; }
	void resize(size_t n) {
		cap = n + 1;
		ptr = Reallocate<wchar_t>(ptr, cap);
	}
	wchar_t& operator[](size_t n) {
		return ptr[n];
	}
	wchar_t operator[](size_t n) const {
		return ptr[n];
	}
	wstring substr(size_t start) const {
		return wstring(ptr + start);
	}
	wstring& operator=(const wstring& other) {
		if (&other != this) {
			cap = other.cap;
			ptr = Reallocate<wchar_t>(ptr, cap);
			CopyString(ptr, other.ptr);
		}
		return *this;
	}
private:
	size_t cap;
	wchar_t* ptr;
};

wstring operator+(const wstring& left, const wstring& right) {
	return wstring(left).append(right);
}

wstring& operator+=(wstring& left, const wstring& right) {
	return left.append(right);
}

// Doesn't handle redirected output
void Write(const wstring& str) {
	auto handle = GetStdHandle(STD_OUTPUT_HANDLE);
	if (handle == INVALID_HANDLE_VALUE || handle == nullptr) return;

	DWORD b;
	WriteConsoleW(handle, str.data(), str.size(), &b, nullptr);
}

void WriteLine(const wstring& str) {
	Write(str + L"\n");
}

wstring Win32ErrorToString(DWORD errCode) {
	LocalPtr<wchar_t> result;

	FormatMessageW(
		FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_IGNORE_INSERTS,
		nullptr,
		errCode,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
		reinterpret_cast<wchar_t*>(&result),
		0,
		nullptr);

	if (result.get() == nullptr) {
		return L"Unknown error.";
	}

	return result.get();
}

template<class T, int Base, unsigned int Exponent>
struct IntPower {
	static constexpr T pow() {
		return IntPower<T, Base, Exponent - 1>::pow() * Base;
	}
};

template<class T, int Base>
struct IntPower<T, Base, 0> {
	static constexpr T pow() {
		return 1;
	}
};

wstring ULongToString(ULONG num) {
	const int digits = 10;
	wstring result;	
	result.resize(digits);
	auto largest = IntPower<ULONG, 10, digits - 1>::pow();
	for (; largest != 0 && (num / largest) % 10 == 0; largest /= 10) {}

	int i = 0;
	for (; largest != 0; largest /= 10, ++i) {
		result[i] = L'0' + ((num / largest) % 10);
	}

	result[i] = L'\0';
	return result;
}

void PrintHResult(HRESULT hres) {
	auto errCode = HResultToWin32Error(hres);
	auto errorText = Win32ErrorToString(errCode);
	WriteLine(errorText.data());
}

wstring StringToLower(const wstring& str) {
	auto len = LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_LOWERCASE, str.data(), str.size() + 1, nullptr, 0, nullptr, nullptr, 0);
	wstring lowerCased;
	lowerCased.resize(len - 1);
	LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_LOWERCASE, str.data(), str.size() + 1, lowerCased.data(), len, nullptr, nullptr, 0);
	return lowerCased;
}

size_t FindPathPrefix(const wstring& path, const wstring& prefix) {
	if (prefix.size() > path.size()) return 0;

	size_t i = 0;

	for (; prefix[i] != 0; ++i) {
		if (path[i] != prefix[i]) {
			return 0;
		}
	}

	if (path[i] != L'\\') return 0;

	return i;
}

wstring GuidifyPath(const wstring& path) {
	const GUID folders[] = {
		FOLDERID_System,
		FOLDERID_ProgramFiles,
		FOLDERID_Windows,
		FOLDERID_ProgramFilesX86
	};

	for (auto& folder : folders) {
		CoTaskPtr<wchar_t> knownFolderPath;
		SHGetKnownFolderPath(folder, KF_FLAG_DEFAULT, nullptr, &knownFolderPath);

		auto lowerKnownFolderPath = StringToLower(knownFolderPath.get());
		auto prefixPos = FindPathPrefix(path, lowerKnownFolderPath);
		if (prefixPos == 0) continue;

		wchar_t guid[40];
		StringFromGUID2(folder, guid, 40);
		auto guidLower = StringToLower(guid);

		return guidLower + path.substr(prefixPos);
	}

	return path;
}

template <typename TClass, typename TInterface>
HRESULT ComCreate(TInterface** out) {
	return CoCreateInstance(__uuidof(TClass), nullptr, CLSCTX_INPROC_SERVER, __uuidof(TInterface), reinterpret_cast<void**>(out));
}

#define CHECK_RESULT(hres, message) do { \
	HRESULT h = (hres); \
	if (h != S_OK) { \
		WriteLine(message); \
		PrintHResult(h); \
		return 1; \
	} \
} while (0)

int wmain(int argc, wchar_t** argv) {
	if (argc != 2) {
		WriteLine(L"Usage: hashlnk PATH");
		return 1;
	}

	CHECK_RESULT(CoInitialize(nullptr), L"Failed to initialize COM.");

	ComPtr<IShellLinkW> link;
	CHECK_RESULT((ComCreate<ShellLink, IShellLinkW>(&link)), L"Failed to create IShellLinkW.");

	ComPtr<IPersistFile> linkFile;
	CHECK_RESULT(link->QueryInterface<IPersistFile>(&linkFile), L"Failed to create IPersistFile.");

	CHECK_RESULT(linkFile->Load(argv[1], STGM_READWRITE), L"Failed to load link file.");

	ComPtr<IPropertyStore> propStore;
	CHECK_RESULT(link->QueryInterface<IPropertyStore>(&propStore), L"Failed to create IPropertyStore.");

	PROPVARIANT pv;
	wstring linkArgs;
	CHECK_RESULT(propStore->GetValue(PKEY_Link_Arguments, &pv), L"Failed to get link arguments.");
	if (pv.pwszVal == nullptr) {
		linkArgs = L"";
	} else {
		linkArgs = StringToLower(pv.pwszVal);
		CoTaskMemFree(pv.pwszVal);
	}

	wchar_t path[MAX_PATH];
	CHECK_RESULT(link->GetPath(path, MAX_PATH, nullptr, 0), L"Failed to get link target.");
	auto lowerPath = StringToLower(path);
	auto linkTargetGuidy = GuidifyPath(lowerPath);

	wstring concat = linkTargetGuidy + linkArgs + L"do not prehash links.  this should only be done by the user.";

	ULONG hash;
	CHECK_RESULT(HashData(reinterpret_cast<BYTE*>(concat.data()), concat.size() * sizeof(wchar_t), reinterpret_cast<BYTE*>(&hash), sizeof(hash)), L"Failed to calculate hash.");
	
	Write(L"Hash calculated as ");
	Write(ULongToString(hash).data());
	Write(L"\n");

	pv.vt = VT_UI4;
	pv.ulVal = hash;
	propStore->SetValue(PKEY_WinX_Hash, pv);
	CHECK_RESULT(propStore->Commit(), L"Failed to commit property store.");

	CHECK_RESULT(linkFile->Save(nullptr, false), L"Failed to save link file");

	WriteLine(L"Link file updated.");

	return 0;
}

extern "C" void __cdecl wmainCRTStartup() {	
	int argc;
	wchar_t** argv = CommandLineToArgvW(GetCommandLineW(), &argc);

	int result = wmain(argc, argv);

	LocalFree(argv);

	ExitProcess(result);
}