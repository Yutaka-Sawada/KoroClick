// Copyright : 2023-04-14 Yutaka Sawada
// License : The MIT license

// Main.cpp : アプリケーションのエントリ ポイントを定義します。
//

// // SDKDDKVer.h をインクルードすると、利用できる最も高いレベルの Windows プラットフォームが定義されます。
// 以前の Windows プラットフォーム用にアプリケーションをビルドする場合は、WinSDKVer.h をインクルードし、
// サポートしたいプラットフォームに _WIN32_WINNT マクロを設定してから SDKDDKVer.h をインクルードします。
#include <SDKDDKVer.h>

#define WIN32_LEAN_AND_MEAN             // Windows ヘッダーからほとんど使用されていない部分を除外する

// Windows ヘッダー ファイル
#include <windows.h>
#include <shellapi.h>

// C ランタイム ヘッダー ファイル
#include <stdlib.h>
#include <malloc.h>
#include <memory.h>
#include <tchar.h>

#include "resource.h"

// 第１６１章 キーボード・フック
// http://www.kumei.ne.jp/c_lang/sdk2/sdk_161.htm

#include "ghook.h"
#pragma comment(lib, "ghook.lib")


// グローバル変数:
WCHAR szWindowClass[] = L"KoroClick";	// メイン ウィンドウ クラス名
HANDLE hMutex;	// 多重起動防止用
BOOL bHook = FALSE;
BOOL bTray = FALSE;

// デバッグ出力用
//#define DEBUG_OUTPUT

#ifdef DEBUG_OUTPUT
#define MAX_LINE_NUMBER 16
#define MAX_LINE_LENGTH 256
#define LINE_POS_LEFT 10
#define LINE_POS_TOP 10
#define LINE_POS_HEIGHT 20
wchar_t multi_line[MAX_LINE_NUMBER][MAX_LINE_LENGTH];
int start_line = 0;
#endif


// タスクトレイ・アイコンの操作
// https://blog.goo.ne.jp/masaki_goo_2006/e/a067535abc8b6f1851db69bcdcf4b761
#define ID_TRAYICON     (1)                 // 複数のアイコンを識別するためのID定数
#define WM_TASKTRAY     (WM_APP + 1)        // タスクトレイのマウス・メッセージ定数

// Put icon on Task Tray
int PutTrayIcon(HWND hWnd)
{
	// 構造体のセット
	NOTIFYICONDATA nid = { 0 };
	nid.cbSize              = sizeof(NOTIFYICONDATA);
	nid.uFlags              = (NIF_ICON | NIF_MESSAGE | NIF_TIP);
	nid.hWnd                = hWnd;				// ウインドウのハンドル
	nid.hIcon               = LoadIcon(GetModuleHandle(NULL), MAKEINTRESOURCE(IDI_PRESS));	// アイコンのハンドル
	nid.uID                 = ID_TRAYICON;		// アイコン識別子の定数
	nid.uCallbackMessage    = WM_TASKTRAY;		// 通知メッセージの定数
	lstrcpy( nid.szTip, TEXT("ころクリック") );	// チップヘルプの文字列

	// ここでアイコンの追加
	if (Shell_NotifyIcon(NIM_ADD, &nid) == FALSE){
#ifdef DEBUG_OUTPUT
		SendMessage(hWnd, WM_SETTEXT, 0, (LPARAM)(L"タスクトレイにアイコン登録失敗"));
#endif
		return 0;	// 登録に失敗した
	}
	return 1;
}
void RemoveTrayIcon(HWND hWnd)
{
	// 構造体のセット
	NOTIFYICONDATA nid = { 0 };
	nid.cbSize              = sizeof(NOTIFYICONDATA);
	nid.hWnd                = hWnd;			// ウインドウのハンドル
	nid.uID                 = ID_TRAYICON;	// アイコン識別子の定数

	// ここでアイコンの削除
	Shell_NotifyIcon(NIM_DELETE, &nid);
}
void ChangeTrayIcon(HWND hWnd, int icon_id)
{
	// 構造体のセット
	NOTIFYICONDATA nid = { 0 };
	nid.cbSize              = sizeof(NOTIFYICONDATA);
	nid.uFlags              = NIF_ICON;
	nid.hWnd                = hWnd;			// ウインドウのハンドル
	nid.hIcon               = LoadIcon(GetModuleHandle(NULL), MAKEINTRESOURCE(icon_id));	// アイコンのハンドル
	nid.uID                 = ID_TRAYICON;	// アイコン識別子の定数

	// ここでアイコンの変更
	Shell_NotifyIcon(NIM_MODIFY, &nid);
}


// このコード モジュールに含まれる関数の宣言を転送します:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

	// Windows上で多重起動を防止する方法
	// https://www.daccho-it.com/program/WinApi/unmultiboot.htm
	hMutex = CreateMutex(NULL, TRUE, L"KoroClickSingleOnly");
    // すでに起動しているか判定
	if (GetLastError() == ERROR_ALREADY_EXISTS){
		// すでに起動している。
		ReleaseMutex(hMutex);
		CloseHandle(hMutex);
		return FALSE;
	}

	// TODO: ここにコードを挿入してください。
	if (!MyRegisterClass(hInstance))
		return FALSE;
	// アプリケーション初期化の実行:
	if (!InitInstance (hInstance, nCmdShow))
		return FALSE;

	MSG msg;

	// メイン メッセージ ループ:
	while (GetMessage(&msg, nullptr, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	// Mutexの開放
	ReleaseMutex(hMutex);
	CloseHandle(hMutex);

	return (int) msg.wParam;
}


//
//  関数: MyRegisterClass()
//
//  目的: ウィンドウ クラスを登録します。
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEXW wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style          = CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc    = WndProc;
	wcex.cbClsExtra     = 0;
	wcex.cbWndExtra     = 0;
	wcex.hInstance      = hInstance;
	wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MAIN));
	wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
	wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName   = L"MYMENU";    // 不要なら NULL にすればいい
	wcex.lpszClassName  = szWindowClass;
	wcex.hIconSm        = NULL;

	return RegisterClassExW(&wcex);
}

//
//   関数: InitInstance(HINSTANCE, int)
//
//   目的: インスタンス ハンドルを保存して、メイン ウィンドウを作成します
//
//   コメント:
//
//        この関数で、グローバル変数でインスタンス ハンドルを保存し、
//        メイン プログラム ウィンドウを作成および表示します。
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	HWND hWnd = CreateWindowW(szWindowClass,
				L"ころクリック",
				WS_OVERLAPPEDWINDOW,
				CW_USEDEFAULT,	// X座標
				CW_USEDEFAULT,	// Y座標
				320,	// 幅
				400,	// 高さ
				nullptr, nullptr, hInstance, nullptr);

	if (!hWnd)
		return FALSE;

	// ウィンドウを表示するかどうかは後で決める
	//ShowWindow(hWnd, nCmdShow);
	//UpdateWindow(hWnd);

	return TRUE;
}

//
//  関数: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  目的: メイン ウィンドウのメッセージを処理します。
//
//  WM_PAINT    - メイン ウィンドウを描画する
//  WM_DESTROY  - 中止メッセージを表示して戻る
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	static HMENU hMenu;

	switch (message){
		case WM_CREATE:
#ifdef DEBUG_OUTPUT
			ZeroMemory(multi_line, sizeof(multi_line));
#endif

			// タスクトレイにアイコンを置く
			bTray = PutTrayIcon(hWnd);
			if (bTray == FALSE){
				// エラーが発生したら操作できなくなるので、ウィンドウを表示する
				ShowWindow(hWnd, SW_SHOWNORMAL);
			}

			// 自動的にフックする
			hMenu = GetMenu(hWnd);
			if (MySetHook(hWnd) != 0){
				// フックに成功した時だけ
				bHook = TRUE;
				CheckMenuItem(hMenu, IDM_HOOK, MF_BYCOMMAND | MF_CHECKED);
			}
			break;
		case WM_COMMAND:
			switch (LOWORD(wParam)){
				case IDM_EXIT:
					SendMessage(hWnd, WM_CLOSE, 0, 0);
					break;
				case IDM_HOOK:
					if (bHook){
						if (MyEndHook() != 0){
							bHook = FALSE;
							CheckMenuItem(hMenu, IDM_HOOK, MF_BYCOMMAND | MF_UNCHECKED);
							if (bTray)
								ChangeTrayIcon(hWnd, IDI_PAUSE);
						}
					} else {
						if (MySetHook(hWnd) != 0){
							bHook = TRUE;
							CheckMenuItem(hMenu, IDM_HOOK, MF_BYCOMMAND | MF_CHECKED);
							if (bTray)
								ChangeTrayIcon(hWnd, IDI_PRESS);
						}
					}
					break;
			}
			break;
		case WM_SYSCOMMAND:
			if (wParam == SC_MINIMIZE){
				if (bTray == FALSE){
					// タスクトレイにアイコンが無い場合は再度置く
					bTray = PutTrayIcon(hWnd);
				}
				if (bTray == TRUE){
					// アイコンを登録出来たら、ウィンドウを隠す
					ShowWindow(hWnd, SW_HIDE);
					return 0;
				} else {
					// アイコンの登録に失敗したら普通に最小化する
					return DefWindowProc(hWnd, message, wParam, lParam);
				}
			} else {
				return DefWindowProc(hWnd, message, wParam, lParam);
			}
			break;
#ifdef DEBUG_OUTPUT
		case WM_SETTEXT:
			{
				// 空の行を探す
				int current_line = 0;
				while (current_line < MAX_LINE_NUMBER){
					if (multi_line[current_line][0] == 0)
						break;
					current_line++;
				}
				// 空の行が無ければ、上にずらす
				if (current_line == MAX_LINE_NUMBER){
					current_line = 1;
					while (current_line < MAX_LINE_NUMBER){
						wcscpy_s(multi_line[current_line - 1], MAX_LINE_LENGTH, multi_line[current_line]);
						current_line++;
					}
					current_line--;
					start_line++;
				}
				wsprintf(multi_line[current_line], L"%3d ", start_line + current_line);
				wcscpy_s(multi_line[current_line] + 4, MAX_LINE_LENGTH - 4, (wchar_t *)lParam);
				InvalidateRect(hWnd, NULL, TRUE); 
			}
			break;
		case WM_PAINT:
			{
				HDC hdc;
				PAINTSTRUCT paint;
				hdc = BeginPaint(hWnd, &paint);
				int current_line = 0;
				while (current_line < MAX_LINE_NUMBER){
					if (multi_line[current_line][0] == 0)
						break;
					TextOut(hdc, LINE_POS_LEFT, LINE_POS_TOP + LINE_POS_HEIGHT * current_line,
							multi_line[current_line], (int)wcslen(multi_line[current_line]));
					current_line++;
				}
				EndPaint(hWnd, &paint);
			}
			break;
#endif
		// タスクトレイをクリックした時のユーザー定義メッセージ
		case WM_TASKTRAY:
			if (wParam == ID_TRAYICON){	// アイコンの識別コード
				switch (lParam){
#ifdef DEBUG_OUTPUT
					case WM_LBUTTONUP:
						// 左ボタンが離されたら、ウィンドウを表示する
						ShowWindow(hWnd, SW_SHOWNORMAL);
						break;
#endif
					case WM_RBUTTONUP:
						// 右ボタンが離されたら、メニューを表示する
						POINT pt;
						if (GetCursorPos(&pt) != 0){
							SetForegroundWindow(hWnd);
							BOOL ret;
							ret = TrackPopupMenu(
									GetSubMenu(hMenu, 0),
									0,
									pt.x,
									pt.y,
									0,
									hWnd,
									NULL
								);
#ifdef DEBUG_OUTPUT
							if (ret == 0){
								SetWindowText(hWnd, L"TrackPopupMenu failed");
							}
						} else {
							SetWindowText(hWnd, L"GetCursorPos failed");
#endif
						}
						break;
					case WM_LBUTTONDBLCLK:
						// 左ダブルクリックされたら、フックを切り替える
						if (bHook){
							if (MyEndHook() != 0){
								bHook = FALSE;
								CheckMenuItem(hMenu, IDM_HOOK, MF_BYCOMMAND | MF_UNCHECKED);
								if (bTray)
									ChangeTrayIcon(hWnd, IDI_PAUSE);
							}
						} else {
							if (MySetHook(hWnd) != 0){
								bHook = TRUE;
								CheckMenuItem(hMenu, IDM_HOOK, MF_BYCOMMAND | MF_CHECKED);
								if (bTray)
									ChangeTrayIcon(hWnd, IDI_PRESS);
							}
						}
						break;
				}
			}
			break;
		case WM_CLOSE:
			if (bHook){
				// 終了前に自動的にフックを解除する
				if (MyEndHook() != 0){
					bHook = FALSE;
				} else {
					MessageBox(hWnd, L"フックが解除されていません", L"注意！！", MB_OK);
				}
			}
			if (bTray){
				// タスクトレイのアイコンを消す
				RemoveTrayIcon(hWnd);
			}

			DestroyWindow(hWnd);
			break;
		case WM_DESTROY:
			PostQuitMessage(0);
			break;
		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

