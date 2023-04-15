// Copyright : 2023-04-15 Yutaka Sawada
// License : The MIT license

// dllmain.cpp : DLL アプリケーションのエントリ ポイントを定義します。
#define WIN32_LEAN_AND_MEAN             // Windows ヘッダーからほとんど使用されていない部分を除外する
// Windows ヘッダー ファイル
#include <windows.h>

/*
第１６１章 キーボード・フック
http://www.kumei.ne.jp/c_lang/sdk2/sdk_161.htm

第１６２章 マウス・フック
http://www.kumei.ne.jp/c_lang/sdk2/sdk_162.htm
*/

#include "ghook.h"

// デバッグ出力用
//#define DEBUG_OUTPUT

// プロセス間で共有されるグローバル変数
#pragma data_seg("MY_DATA")
HWND hWnd = NULL;	// メイン・ウィンドウのハンドル
volatile int button_push = 0;
volatile UINT_PTR timer_id = 0;	// タイマーの識別番号
#pragma data_seg()
// 複数の DLL が起動しても、一度に応答するのは一個だけのはず。
// だから、Mutex を使わずとも、volatile だけで十分な気がする。

// グローバル変数
HINSTANCE hInst;
HHOOK hMyHook;


// ボタンを離すのが速すぎると、押したことを認識してくれない。
// 最低でも 32~48 ms 以上にすること。遅いゲーム向けに 96 ms にしておく。
// VK_LBUTTON をチェックする頻度が毎秒 10回 以上なら、取りこぼししないはず。
// しかし、200 ms とか 500 ms まで伸ばしても、反応しないことがある。
#define TIMER_WAIT	96

// マウスボタンを離す
void ReleaseButton()
{
	INPUT inputs[1];
	ZeroMemory(inputs, sizeof(inputs));

	// Virtual-Key mouse left button is released
	inputs[0].type = INPUT_KEYBOARD;
	inputs[0].ki.wVk = VK_LBUTTON;
	inputs[0].ki.dwFlags = KEYEVENTF_KEYUP;

	SendInput(ARRAYSIZE(inputs), inputs, sizeof(INPUT));

	// ホイール高速回転で連続クリックしても大丈夫なよう、ボタンを離してから 0 にする。
	button_push = 0;
	timer_id = 0;
}

// タイマーが呼び出す関数
// http://wisdom.sakura.ne.jp/system/winapi/win32/win47.html
void CALLBACK TimerProc(HWND hwnd , UINT uMsg ,UINT idEvent , DWORD dwTime)
{
	// 一度だけ実行するので、すぐにタイマーを消す。
	KillTimer(hwnd, idEvent);

	// マウスボタンを離す
	ReleaseButton();
}

// マウスボタンを押す
int PressButton()
{
	// タイマーイベント待ちなら、タイマーを消してから、ボタンを離す。
	if (button_push != 0){
		KillTimer(NULL, timer_id);
		ReleaseButton();
	}
	// ボタンを押した状態でさらに押すと不具合が発生するから、既に押されてたら終わる。
	// ２重押しは本来あり得ない動作なので、最悪の場合アプリケーションがフリーズする。
	if (GetAsyncKeyState(VK_LBUTTON) & 0x8000){
#ifdef DEBUG_OUTPUT
		SendMessage(hWnd, WM_SETTEXT, 0, (LPARAM)(L"GetAsyncKeyState(VK_LBUTTON)"));
#endif
		return 0;
	}

	// 先にマウスイベントを送ることでウィンドウをアクティブにする。
	// キーコードをスキャンする頻度によっては動かないので、後からキーコードも送る。
	// なお、MOUSEEVENTF_LEFTDOWN を送信した時点で、WM_LBUTTONDOWN が送られて、
	// 同時に VK_LBUTTON もセットされるみたい。
	// 逆に、INPUT_KEYBOARD で VK_LBUTTON をセットしても、WM_LBUTTONDOWN は送られない。
	INPUT inputs[3];
	ZeroMemory(inputs, sizeof(inputs));

	// Mouse left button down
	inputs[0].type = INPUT_MOUSE;
	inputs[0].mi.dwFlags = MOUSEEVENTF_LEFTDOWN;

	// Mouse left button up
	inputs[1].type = INPUT_MOUSE;
	inputs[1].mi.dwFlags = MOUSEEVENTF_LEFTUP;

	// Virtual-Key mouse left button is pressed
	inputs[2].type = INPUT_KEYBOARD;
	inputs[2].ki.wVk = VK_LBUTTON;

	if (SendInput(ARRAYSIZE(inputs), inputs, sizeof(INPUT)) != ARRAYSIZE(inputs)){
#ifdef DEBUG_OUTPUT
		SendMessage(hWnd, WM_SETTEXT, 0, (LPARAM)(L"SendInput failed"));
#endif
		return 1;
	}

	// マウスボタンを押した後、一定時間経過後にボタンを離す。
	// Sleep() 関数だと反応しないアプリケーションが存在するので、タイマーを使う。
	timer_id = SetTimer(NULL, 0, TIMER_WAIT, (TIMERPROC)TimerProc);
	if (timer_id != 0){
		// ボタンを押した後に、離すのを待ってる印
		button_push = 1;

	} else {
#ifdef DEBUG_OUTPUT
		wchar_t str[256];
		wsprintf(str, L"SetTimer failed = %u", GetLastError());
		SendMessage(hWnd, WM_SETTEXT, 0, (LPARAM)str);
#endif

		// タイマー作成に失敗した場合はすぐにボタンを離す
		ReleaseButton();
	}

	return 0;
}

BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD fdReason, PVOID pvReserved)
{
	switch (fdReason)
	{
		case DLL_PROCESS_ATTACH:
			hInst = hInstance;
#ifdef DEBUG_OUTPUT
			if (hWnd){
				SendMessage(hWnd, WM_SETTEXT, 0, (LPARAM)(L"プロセスにアタッチしました"));
			} else {
				//MessageBox(NULL, L"プロセスにアタッチしました", L"アタッチ", MB_OK);
			}
#endif
			break;
		case DLL_PROCESS_DETACH:
#ifdef DEBUG_OUTPUT
			if (hWnd){
				SendMessage(hWnd, WM_SETTEXT, 0, (LPARAM)(L"プロセスからデタッチしました"));
			} else {
				//MessageBox(NULL, L"プロセスからデタッチしました", L"デタッチ", MB_OK);
			}
#endif
			break;
/*
		case DLL_THREAD_ATTACH:
#ifdef DEBUG_OUTPUT
			if (hWnd){
				SendMessage(hWnd, WM_SETTEXT, 0, (LPARAM)(L"スレッドにアタッチしました"));
			} else {
				//MessageBox(NULL, L"スレッドにアタッチしました", L"アタッチ", MB_OK);
			}
#endif
			break;
		case DLL_THREAD_DETACH:
#ifdef DEBUG_OUTPUT
			if (hWnd){
				SendMessage(hWnd, WM_SETTEXT, 0, (LPARAM)(L"スレッドからデタッチしました"));
			} else {
				//MessageBox(NULL, L"スレッドからデタッチしました", L"デタッチ", MB_OK);
			}
#endif
			break;
*/
	}

	return TRUE;
}

// 64-bit OS で動かす場合でも、32-bit DLL を使わないといけないっぽい。
// https://learn.microsoft.com/ja-jp/windows/win32/api/winuser/nf-winuser-setwindowshookexw
int MySetHook(HWND hMainWindow)
{
	// 呼び出し元ウィンドウのハンドルを記録しておく
	hWnd = hMainWindow;

	hMyHook = SetWindowsHookEx(WH_MOUSE, (HOOKPROC)MyHookProc, hInst, 0);
	if (hMyHook != NULL){
#ifdef DEBUG_OUTPUT
		SendMessage(hWnd, WM_SETTEXT, 0, (LPARAM)(L"フック成功"));
#endif
		// 初期状態では押されてない
		button_push = 0;
		return 1;
	} else {
#ifdef DEBUG_OUTPUT
		SendMessage(hWnd, WM_SETTEXT, 0, (LPARAM)(L"フック失敗"));
#endif
		return 0;
	}
}

int MyEndHook(void)
{
	// ボタンが押されたままなら離す
	if (button_push != 0){
		KillTimer(NULL, timer_id);
		ReleaseButton();
	}

	if (UnhookWindowsHookEx(hMyHook) != 0){
#ifdef DEBUG_OUTPUT
		SendMessage(hWnd, WM_SETTEXT, 0, (LPARAM)(L"フック解除成功"));
#endif
		return 1;
	} else {
#ifdef DEBUG_OUTPUT
		SendMessage(hWnd, WM_SETTEXT, 0, (LPARAM)(L"フック解除失敗"));
#endif
		return 0;
	}
}

LRESULT CALLBACK MyHookProc(int nCode, WPARAM wp, LPARAM lp)
{
	if (nCode < 0)
		return CallNextHookEx(NULL, nCode, wp, lp);

	MOUSEHOOKSTRUCT *pmh;
	pmh = (MOUSEHOOKSTRUCT *)lp;
	// マウスホイールを回転させるか、ホイールを押した場合
	if ( (wp == WM_MOUSEWHEEL) || (wp == WM_MBUTTONDOWN) || (wp == WM_NCMBUTTONDOWN) ){
		// LONG pt.x	カーソルの x 座標と y 座標 (画面座標)
		// LONG pt.y
		// HWND hwnd	マウス イベントに対応するマウス メッセージを受け取るウィンドウへのハンドル
		// UINT wHitTestCode	ヒット テスト値 https://learn.microsoft.com/ja-jp/windows/win32/inputdev/wm-nchittest
		//wchar_t str[256];
		//wsprintf(str, L"X = %d, Y = %d, hwnd = %p", pmh->pt.x, pmh->pt.y, pmh->hwnd);
		//MessageBox(NULL, str, L"MyHookProc", MB_OK);

		// マウスの左クリックをエミュレートする
		if (PressButton() != 0)
			return CallNextHookEx(NULL, nCode, wp, lp);
		return TRUE;

	// ホイールを押した後に離しても反応しないようにする
	} else if ( (wp == WM_MBUTTONUP) || (wp == WM_NCMBUTTONUP) ){
		return TRUE;
	}

	// 第1引数は無視されるらしい。hMyHook を DLL 間で共有する必要性は無い。
	// https://learn.microsoft.com/ja-jp/windows/win32/api/winuser/nf-winuser-callnexthookex
	return CallNextHookEx(NULL, nCode, wp, lp);
}

