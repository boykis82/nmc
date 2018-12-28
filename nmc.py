from time import sleep
from nateonmemo import NateonMemo
from nateonmemo import NateonMemoAutomation
import logging
import argparse
from datetime import datetime, timedelta

SEND_MEMO_PATH = 'd:\\memo\\send'
RECV_MEMO_PATH = 'd:\\memo\\recv'
#TEMP_PATH = 'd:\\memo\\temp'

if __name__ == '__main__':
    # logging 모듈 setup
    log = logging.getLogger("nmc")
    log.setLevel(logging.INFO)

    formatter = logging.Formatter('[%(asctime)s][%(levelname)s] %(message)s', '%Y-%m-%d %H:%M:%S')
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    log.addHandler(stream_handler)

    # default argument setup
    yesterday = (datetime.today() - timedelta(1)).strftime('%Y%m%d')
    one_mth_ago = (datetime.today() - timedelta(30)).strftime('%Y%m%d')    

    # argument parser
    arg_parser = argparse.ArgumentParser(description='nmc parameters')
    arg_parser.add_argument("--from_date", help="from date : YYYYMMDD", type=str, default=one_mth_ago)
    arg_parser.add_argument("--to_date", help="to date : YYYYMMDD", type=str, default=yesterday)    
    arg_parser.add_argument("--save_recv", help="save recv box? : Y/N", type=str, default='Y')
    arg_parser.add_argument("--save_send", help="save send box? : Y/N", type=str, default='Y')
    arg_parser.add_argument("--save_all", help="save all? : Y: all /N: from current cursor", type=str, default='Y')
    arg_parser.add_argument("--person", help="person", type=str, default='')
    arg_parser.add_argument("--search_word", help="search_word", type=str, default='')
    args = arg_parser.parse_args()    

    log.info(f'from = {args.from_date}, to = {args.to_date}, save_recv = {args.save_recv}, \
        save_send = {args.save_send}, save_all = {args.save_all}, person={args.person}, search_word={args.search_word}')

    log.info('auto saver initializing...')

    # NateonMemoAutomation object 생성 & bizmemobox와 연결
    nma = NateonMemoAutomation(
        RECV_MEMO_PATH, SEND_MEMO_PATH)
    nma.initialize()

    log.info('auto saver initialize completed!!')

    sleep(1)

    log.info('auto saver set search conditions!!')

    # 조회 조건 설정
    nma.set_search_conditions(
        from_date=args.from_date, to_date=args.to_date, check_period=True, send_recv_person=args.person, search_word=args.search_word)

    log.info('start saving memo!!')

    
    # 받은 쪽지함 저장여부=Y
    if args.save_recv == 'Y':
        nma.select_recv_box_tree()
        if args.save_all == 'Y':
            nma.click_search_btn()
            nma.move_to_first_memo_pos()
        nma.save_recv_memo_box()

    # 보낸 쪽지함 저장여부=Y
    if args.save_send == 'Y':
        nma.select_send_box_tree()
        nma.click_search_btn()
        if args.save_all == 'Y':
            nma.click_search_btn()      
            nma.move_to_first_memo_pos()  
        nma.save_send_memo_box()
    

    sleep(1)

    log.info('saving memo completed!!')