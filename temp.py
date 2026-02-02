@autofuncmgr.AutoFuncTinyDecorator
def run(script_path: str, target: str = "", target_area: str = "", display: Display = Display.FRONT):
    """
    Args:
        target (str): 검증 대상
        target_area (str): 검증 대상이 위치한 화면 상의 영역
        display (Display): Display.FRONT: 전석, Display.REAR: 후석

    Returns:
        ret (int): 함수 실행 결과
        msg (str): 함수 실행 결과 메시지
    """
    
    script_name = Path(script_path).stem
    
    # TMP 이미지 갱신
    if display == Display.FRONT:
        # tiny.CaptureScreen.run() # CVM 검증 대상 탐색
        if not avnctrl.capture_screen():
            report.add_step_error_finalize(base.RETCODE.AVN_CAPTURE_ERROR)
            return base.RETCODE.AVN_CAPTURE_ERROR, False
    else:
        # tiny.CaptureScreenOnRear.run() # CVM 검증 대상 탐색
        if not avnctrl.capture_screen_rear():
            report.add_step_error_finalize(base.RETCODE.AVN_CAPTURE_ERROR)
            return base.RETCODE.AVN_CAPTURE_ERROR, False
        
    report.add_step_prepare(target_area, target, do_capture=True)

    target_list = []

    

    return ret
