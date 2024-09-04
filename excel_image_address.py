import pandas as pd
import xml.etree.ElementTree as ET
import os, shutil


def unpack_xlsx(src: str) -> str:
    """
    拡張子を .xlsx から .zip に変更して解凍する。
    エクセルファイルと同階層に展開する。

    Parameters
    ----------
    src : str
        エクセルファイル名

    Returns
    -------
    upk : str
        解凍後のフォルダ名
    """
    # 拡張子の変更
    dst = src.replace('.xlsx', '.zip')  # コピー先のファイル名
    shutil.copyfile(src, dst)           # 拡張子を .zip に変更して同階層にコピー

    # 解凍
    upk = os.path.splitext(src)[0]      # エクセルファイルと同名の解凍後フォルダ名
    shutil.unpack_archive(dst, upk)     # .zip を解凍

    return upk


def parse_xml(path: str) -> pd.DataFrame:
    """
    シート情報をもつ drawing[x].xml を解析する。
    画像が配置されたセル番地と画像IDの対応関係を取得する。

    Parameters
    ----------
    path : str
        drawing[x].xml のパス

    Returns
    -------
    df : pd.DataFrame
        セル番地と画像IDの対応関係
    """

    # タグの名前空間（エクセルファイルに応じて変更してください。）
    xdr = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

    tree = ET.parse(path)   # xmlのパーサ
    root = tree.getroot()   # root として <xdr:wsDr> タグの部分を取得

    col_elems = root.findall(f'.//{xdr}twoCellAnchor/{xdr}from/{xdr}col')   # 画像位置の列番号の要素リストを取得
    cols = [int(elem.text) for elem in col_elems]                           # 値を整数に変換

    row_elems = root.findall(f'.//{xdr}twoCellAnchor/{xdr}from/{xdr}row')   # 画像位置の行番号の要素リストを取得
    rows = [int(elem.text) for elem in row_elems]                           # 値を整数に変換

    rid_elems = root.findall(f'.//{xdr}twoCellAnchor/{xdr}pic/{xdr}blipFill/{a}blip')   # rIDの要素リストを取得
    rids = [elem.get(f'{r}embed') for elem in rid_elems]                                # rIDの値を取得

    df = pd.DataFrame({'row': rows, 'col': cols, 'rid': rids})          # (行番号, 列番号, rID) 列をもつ DataFrame を作成
    df = df.sort_values(['row', 'col']).reset_index(drop=True)          # セル番地でソート

    return df


def parse_rel(path: str) -> pd.DataFrame:
    """
    画像のリレーション情報をもつ drawing[x].xml.rels を解析する。
    画像IDと画像ファイル名の対応関係を取得する。

    Parameters
    ----------
    path : str
        drawing[x].xml.rels のパス

    Returns
    -------
    df : pd.DataFrame
        画像IDと画像ファイル名の対応関係
    """

    # タグの名前空間（エクセルファイルに応じて変更してください。）
    xmlns = "{http://schemas.openxmlformats.org/package/2006/relationships}"

    tree = ET.parse(path)   # xmlのパーサ
    root = tree.getroot()   # root として <Relationships> タグの部分を取得

    rel_elems = root.findall(f'.//{xmlns}Relationship')                     # 対応関係の要素リストを取得
    rids = [elem.get('Id') for elem in rel_elems]                           # rIDのリストを取得
    imgs = [os.path.basename(elem.get('Target')) for elem in rel_elems]     # ファイル名のリストを取得

    df = pd.DataFrame({'rid': rids, 'name': imgs})  # (rID, 画像ファイル名) 列をもつ DataFrame を作成

    return df



if __name__ == '__main__':
    
    src = './irasutoya_athletics.xlsx'    # エクセルファイル名
    upk = unpack_xlsx(src)                # 解凍を実行し、解凍先フォルダを取得 

    xml_path = os.path.join(upk, 'xl/drawings/drawing1.xml')            # drawings[x].xml のパス
    rel_path = os.path.join(upk, 'xl/drawings/_rels/drawing1.xml.rels') # drawings[x].xml.rels のパス

    xml_df = parse_xml(xml_path)    # セル番地と画像IDの対応関係
    rel_df = parse_rel(rel_path)    # 画像IDと画像ファイル名の対応関係

    res_df = pd.merge(left=xml_df, right=rel_df, how='inner', on='rid') # 画像IDを照合し、セル番地と画像ファイル名の対応関係を取得
    path = f'{os.path.splitext(src)[0]}_imrels.csv'                     # エクセルファイル名に接尾辞をつける
    res_df.to_csv(path)                                                 # 結果を csv ファイルに出力