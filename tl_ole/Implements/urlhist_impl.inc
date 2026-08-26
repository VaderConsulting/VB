[
	odl,
	uuid(3C374A41-BAE4-11CF-BF7D-00AA006946EE),
]
interface IUrlHistoryStg : stdole.IUnknown {

	[helpstring("Adds a new history entry")]
	HRESULT AddUrl (
		[in] LONG pocsUrl,
		[in] LONG pocsTitle,
		[in] STATURLFLAGS dwFlags);

	[helpstring("Deletes an entry by its URL")]
	HRESULT DeleteUrl(
		[in] LONG pocsUrl,
		[in] LONG dwFlags);

	[helpstring("Returns a STATURL for a given URL")]
	HRESULT QueryUrl (
		[in] LONG pocsUrl,
		[in] LONG dwFlags,
		[in, out] STATURL *lpSTATURL);

	[helpstring("Binds to an object")]	
	HRESULT BindToObject (
		[in] LONG pocsUrl,
		[in, out] UUID *riid,
		[in] long ppvOut);

	[helpstring("Returns an enumerator for URLs")]
	HRESULT EnumUrls (
		[out, retval] IEnumSTATURL **ppEnum);

}
