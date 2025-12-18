
//---------------------------------------------------------------------------
/****************************************************************************
* Name ...........HDACS DLL
* Parameter.......
* Author .........Edgar Hu (humingfei@hotmail.com,13701974214)
* Date ...........2006/02/01
* Company ........HUNDURE TECHNOLOGY CO.,LTD.USA
****************************************************************************/

#ifdef HDACS_EXPORTS
#define HDACS_API __declspec(dllexport)
#else
#define HDACS_API __declspec(dllimport)
#endif


//Reserve for multithread
class HDACS_API CHDACS {
public:
	CHDACS(void);
};

extern HDACS_API int nHDACS;
HDACS_API int fnHDACS(void);

#ifdef __cplusplus
extern "C" {
#endif

	HDACS_API int __stdcall hdacsOpenChannel(HANDLE *hComm,char *sComm,unsigned int iPort);
	HDACS_API int __stdcall hdacsCloseChannel(HANDLE hComm);
	HDACS_API int __stdcall hdacsReadData(HANDLE hComm,unsigned char *cBuffer,int *iDataLen,unsigned int iTimeout);
	HDACS_API int __stdcall hdacsWriteData(HANDLE hComm,unsigned char *cBuffer,int iDataLen,int *iWrittenLen,unsigned int iTimeout);
	HDACS_API int __stdcall hdacsClearBuffer(HANDLE hComm);

#ifdef __cplusplus
}
#endif

