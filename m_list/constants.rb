# frozen_string_literal: true
# M List - 定数定義

module MList
  EXT_NAME = "M List"
  DICT = "m_list"
  CSV_PATH_KEY = "csv_path"
  PROJECT_FOLDER_KEY = "project_folder"
  PROJECT_BASE_NAME_KEY = "project_base_name"
  LIST_IDS_KEY = "list_ids"

  # エンティティ属性キー
  ATTR_MATERIAL_ID = "material_id"
  ATTR_FINISH_ID = "finish_id"
  ATTR_MATERIAL_NAME = "material_name"
  ATTR_PRICE = "price"
  ATTR_NAME = "name"
  ATTR_MEMO = "memo"
  ATTR_SHAPE_TYPE = "shape_type"
  ATTR_HIDDEN = "hidden"

  # 形状判定の閾値
  BOARD_THICKNESS_MAX_MM = 60
  BOARD_MIN_DIM_MM = 80
  BOARD_THICKNESS_RATIO_MAX = 0.5
  BEAM_MIN_D_MM = 5
  BEAM_ASPECT_RATIO_MIN = 2.5
  FREEFORM_MIN_FACE_COUNT = 6
end
