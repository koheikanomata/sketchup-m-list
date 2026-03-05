# frozen_string_literal: true
# M List - 見積もりプラグイン（自動リスト取得・材料割り振り・価格計算）
# Version 1.1.5

require 'sketchup.rb'
require 'json'
require 'csv'
require 'fileutils'
require 'set'

# Phase 4: コード分割（将来のリファクタ用 - 現状はメインファイルで完結）
# require_relative 'm_list/constants'
# require_relative 'm_list/entity_utils'
# require_relative 'm_list/shape_classifier'
# require_relative 'm_list/materials'
# require_relative 'm_list/price_calculator'
# require_relative 'm_list/list_manager'
# require_relative 'm_list/csv_handler'

module MList
  VERSION = "1.1.5"
  EXT_NAME = "M List"
  DICT = "m_list"

  # オンラインアップデート用
  UPDATE_CHECK_URL = "https://raw.githubusercontent.com/koheikanomata/sketchup-m-list/main/version.json"
  UPDATE_DOWNLOAD_URL = "https://github.com/koheikanomata/sketchup-m-list/releases/latest"
  UPDATE_RAW_BASE = "https://raw.githubusercontent.com/koheikanomata/sketchup-m-list/main/"
  # 自動チェック間隔（秒）2週間 = 1209600
  UPDATE_CHECK_INTERVAL = 1_209_600
  UPDATE_CHECK_PREF_KEY = "m_list_last_update_check"
  CSV_PATH_KEY = "csv_path"
  PROJECT_FOLDER_KEY = "project_folder"
  PROJECT_BASE_NAME_KEY = "project_base_name"
  LIST_IDS_KEY = "list_ids"
  ATTR_MATERIAL_ID = "material_id"
  ATTR_FINISH_ID = "finish_id"
  ATTR_THICKNESS_MM = "thickness_mm"
  ATTR_MATERIAL_NAME = "material_name"
  ATTR_PRICE = "price"
  ATTR_NAME = "name"
  ATTR_MEMO = "memo"
  ATTR_SHAPE_TYPE = "shape_type"
  ATTR_HIDDEN = "hidden"
  ATTR_PRICE_OVERRIDE = "price_override"
  ATTR_LAST_CALC_BOUNDS = "last_calc_bounds"
  BOARD_THICKNESS_MAX_MM = 60
  BOARD_MIN_DIM_MM = 80
  BOARD_THICKNESS_RATIO_MAX = 0.5
  BEAM_MIN_D_MM = 5
  BEAM_ASPECT_RATIO_MIN = 2.5
  FREEFORM_MIN_FACE_COUNT = 6

  def self.plugin_root
    File.join(File.dirname(__FILE__), "m_list")
  end

  def self.model
    Sketchup.active_model
  end

  def self.get_selected_objects
    return [] unless model
    model.selection.select { |e| e.is_a?(Sketchup::Group) || e.is_a?(Sketchup::ComponentInstance) }
  end

  def self.entity_type_label(e)
    case e
    when Sketchup::Group then "グループ"
    when Sketchup::ComponentInstance then "コンポーネント"
    when Sketchup::Face then "面"
    when Sketchup::Edge then "線"
    else "その他"
    end
  end

  def self.entity_tag(e)
    return "" unless e.respond_to?(:layer)
    layer = e.layer
    return "" unless layer
    name = (layer.respond_to?(:display_name) ? layer.display_name : nil) || layer.name.to_s
    name = "" if name.to_s.strip.empty?
    (name == "Layer0" || name.downcase == "layer 0") ? "Untagged" : name
  end

  def self.entity_name(e)
    case e
    when Sketchup::ComponentInstance
      name = e.name.to_s
      name = e.definition.name.to_s if name.empty?
      name.empty? ? "(無名コンポーネント)" : name
    when Sketchup::Group
      name = e.name.to_s
      name.empty? ? "Group" : name
    when Sketchup::Face, Sketchup::Edge
      nil
    else
      ""
    end
  end

  def self.size_string(e)
    return nil unless e.respond_to?(:bounds)
    bb = e.bounds
    w = bb.width.to_mm.round
    h = bb.height.to_mm.round
    d = bb.depth.to_mm.round
    "#{w} x #{h} x #{d} mm"
  end

  # bounds の署名文字列（形状変化検出用）。形式: "幅,高さ,奥行"（mm、昇順ソート）
  def self.bounds_signature(entity)
    return nil unless entity&.respond_to?(:bounds)
    dims = [entity.bounds.width.to_mm, entity.bounds.height.to_mm, entity.bounds.depth.to_mm].sort
    return nil if dims[0] <= 0 || dims[1] <= 0 || dims[2] <= 0
    "#{dims[0].round},#{dims[1].round},#{dims[2].round}"
  end

  def self.list_items
    @list_items ||= []
  end

  def self.clear_list
    @list_items = []
  end

  def self.persist_list_ids
    return unless model
    ids = list_items.map { |it| it[:id].to_s }
    model.set_attribute(DICT, LIST_IDS_KEY, ids.to_json)
  end

  # モデルから全 Group/Component を自動取得してリストを構築（CSV は参照しない）
  def self.refresh_list_from_model
    return unless model
    @list_items = []
    model.entities.each do |e|
      next unless e.is_a?(Sketchup::Group) || e.is_a?(Sketchup::ComponentInstance)
      item = build_list_item_tree(e)
      merge_for_update(nil, item, nil, e)
      @list_items << item
    end
    propagate_nested_prices!(@list_items)
    refresh_panel
    show_panel_status("リストを更新しました: #{@list_items.size}件")
  rescue StandardError => e
    UI.messagebox("リスト更新エラー\n#{format_error(e, 'refresh_list_from_model')}")
  end

  # ---------- CSV パス ----------
  def self.csv_path
    @csv_path
  end

  def self.csv_path=(path)
    @csv_path = path
  end

  def self.project_folder_path
    return nil unless model
    model.get_attribute(DICT, PROJECT_FOLDER_KEY, "").to_s.strip
  end

  def self.project_folder_path=(path)
    return unless model && path && !path.to_s.strip.empty?
    model.set_attribute(DICT, PROJECT_FOLDER_KEY, path.to_s)
  end

  def self.project_folder_csv_path
    return nil unless model
    pf = project_folder_path
    return nil if pf.nil? || pf.empty?
    base = model.get_attribute(DICT, PROJECT_BASE_NAME_KEY, "").to_s.strip
    base = File.basename(pf) if base.empty?
    return nil if base.empty?
    File.join(pf, "#{base}_estimate.csv")
  end

  def self.convention_csv_path
    return nil unless model
    mp = model.path.to_s
    return nil if mp.empty?
    dir = File.dirname(mp)
    base = File.basename(mp, ".*")
    return nil if base.empty?
    File.join(dir, "#{base}_estimate.csv")
  end

  def self.resolved_csv_path
    return nil unless model
    conv = convention_csv_path
    return conv if conv
    pf_csv = project_folder_csv_path
    return pf_csv if pf_csv
    nil
  end

  def self.persist_csv_path(path)
    return unless model && path && !path.to_s.strip.empty?
    model.set_attribute(DICT, CSV_PATH_KEY, path.to_s)
    @csv_path = path.to_s
  end

  # ---------- 材料マスタ（面状・棒状・仕上材の3種類）----------
  # プロジェクトフォルダ内を優先、なければプラグイン内のデフォルトを使用
  def self.path_materials_board(folder = nil)
    f = folder || project_folder_path
    if f && !f.to_s.strip.empty?
      p = File.join(f, "materials_board.csv")
      return p if File.exist?(p)
    end
    File.join(plugin_root, "materials_board.csv")
  end

  def self.path_materials_beam(folder = nil)
    f = folder || project_folder_path
    if f && !f.to_s.strip.empty?
      p = File.join(f, "materials_beam.csv")
      return p if File.exist?(p)
    end
    File.join(plugin_root, "materials_beam.csv")
  end

  def self.path_materials_finish(folder = nil)
    f = folder || project_folder_path
    if f && !f.to_s.strip.empty?
      p = File.join(f, "materials_finish.csv")
      return p if File.exist?(p)
    end
    File.join(plugin_root, "materials_finish.csv")
  end

  def self.load_materials_board(folder = nil)
    path = path_materials_board(folder)
    return [] unless File.exist?(path)
    CSV.read(path, headers: true, encoding: "UTF-8").map(&:to_h)
  rescue StandardError
    []
  end

  def self.load_materials_beam(folder = nil)
    path = path_materials_beam(folder)
    return [] unless File.exist?(path)
    CSV.read(path, headers: true, encoding: "UTF-8").map(&:to_h)
  rescue StandardError
    []
  end

  def self.load_materials_finish(folder = nil)
    path = path_materials_finish(folder)
    return [] unless File.exist?(path)
    CSV.read(path, headers: true, encoding: "UTF-8").map(&:to_h)
  rescue StandardError
    []
  end

  # 素材キャッシュをクリア（プロジェクトフォルダの最新CSVを再読み込みする前に呼ぶ）
  def self.clear_materials_cache
    @raw_by_id = nil
    @finish_by_id = nil
  end

  # プロジェクトフォルダ内の素材CSVを再読み込み
  def self.reload_materials_from_project_folder
    return unless model
    pf = project_folder_path
    unless pf && Dir.exist?(pf)
      UI.messagebox("プロジェクトフォルダが設定されていません。\n\n先に「プロジェクト保存」を行ってください。")
      return
    end
    clear_materials_cache
    refresh_panel
    show_panel_status("素材CSVを更新しました: #{pf}")
  end

  # プロジェクトフォルダに素材CSVを保存（プロジェクト保存時に呼ぶ）
  # 既にプロジェクトフォルダにあればそれを、なければプラグイン内のデフォルトをコピー
  def self.write_materials_to_project_folder(folder)
    return unless folder && Dir.exist?(folder)
    [[:beam, :load_materials_beam], [:board, :load_materials_board], [:finish, :load_materials_finish]].each do |kind, loader|
      data = send(loader, folder)
      data = send(loader) if data.nil? || data.empty?
      next if data.nil? || data.empty?
      dst = File.join(folder, "materials_#{kind}.csv")
      headers = data.first.keys
      CSV.open(dst, "w", encoding: "UTF-8", write_headers: true, headers: headers) do |csv|
        data.each { |row| csv << headers.map { |h| row[h] } }
      end
    end
  rescue StandardError => e
    UI.messagebox("素材CSV保存エラー\n#{e.message}")
  end

  # 面状・棒状を統合したID検索用（価格計算で使用）
  def self.raw_by_id
    @raw_by_id ||= begin
      h = {}
      load_materials_board.each { |r| h[r["material_id"].to_s.strip] = r if r["material_id"] }
      load_materials_beam.each { |r| h[r["material_id"].to_s.strip] = r if r["material_id"] }
      h
    end
  end

  def self.finish_by_id
    @finish_by_id ||= load_materials_finish.each_with_object({}) { |r, h| h[r["finish_id"].to_s.strip] = r if r["finish_id"] }
  end

  def self.raw_beam_options
    raw_by_id.select { |_id, r| linear_unit?(r["unit_size"]) }.values
  end

  def self.raw_board_options
    raw_by_id.select { |_id, r| sheet_unit?(r["unit_size"]) }.values
  end

  # 板材CSV行から厚み(mm)を取得。thickness_mm列優先、なければnameからパース
  def self.parse_board_thickness_mm(row)
    return nil unless row
    t = row["thickness_mm"].to_s.strip
    return t.to_f if t != "" && t.to_f > 0
    name = (row["name"] || "").to_s
    m = name.match(/(\d+\.?\d*)\s*mm/i)
    m ? m[1].to_f : nil
  end

  # 板材で利用可能な厚みの一覧（ソート済み）
  def self.board_thickness_options
    raw_board_options.map { |r| parse_board_thickness_mm(r) }.compact.uniq.sort
  end

  # unit_size が 1220x2440 系か（優先規格）
  def self.board_prefer_1220x2440?(row)
    sz = (row["unit_size"] || "").to_s
    sz.match?(/1220\s*[xX×]\s*2440|2440\s*[xX×]\s*1220/)
  end

  # 指定厚みの板材のみ抽出（1220x2440 を優先して先頭に）
  def self.board_options_by_thickness(thickness_mm)
    return [] if thickness_mm.nil? || thickness_mm.to_s.strip.empty?
    t = thickness_mm.to_f
    return [] if t <= 0
    raw_board_options
      .select { |r| (parse_board_thickness_mm(r) - t).abs < 0.01 }
      .sort_by { |r| board_prefer_1220x2440?(r) ? 0 : 1 }
  end

  # ダイナミックコンポーネントの panel_thickness を取得（cm → mm）。無ければ nil
  def self.dc_panel_thickness_mm(entity, parent_entity = nil)
    val = (entity.get_attribute("dynamic_attributes", "panel_thickness", nil) rescue nil)
    val = (parent_entity.get_attribute("dynamic_attributes", "panel_thickness", nil) rescue nil) if (val.nil? || val.to_f <= 0) && parent_entity
    (val && val.to_f > 0) ? (val.to_f * 10) : nil  # cm → mm
  end

  # 板材エンティティの厚み(mm)。DC の panel_thickness を優先、なければ bounds の最小寸法
  def self.board_thickness_mm(entity, parent_entity = nil)
    dc_mm = dc_panel_thickness_mm(entity, parent_entity)
    return dc_mm if dc_mm && dc_mm > 0
    return nil unless entity&.respond_to?(:bounds)
    bb = entity.bounds
    dims = [bb.width.to_mm, bb.height.to_mm, bb.depth.to_mm].sort
    dims[0] > 0 ? dims[0] : nil
  end

  # 厚みが最も近い板材の { material_id, thickness_mm } を返す（同一厚み優先、1220x2440 優先）
  def self.find_best_board_match_by_thickness(entity, parent_entity = nil)
    target_mm = board_thickness_mm(entity, parent_entity)
    return nil unless target_mm && target_mm > 0
    candidates = raw_board_options.map do |r|
      row_mm = parse_board_thickness_mm(r)
      next nil unless row_mm && row_mm > 0
      diff = (row_mm - target_mm).abs
      prefer = board_prefer_1220x2440?(r) ? 0 : 1
      { material_id: r["material_id"].to_s.strip, thickness_mm: row_mm, diff: diff, exact: (diff < 0.01), prefer_1220: prefer }
    end.compact
    return nil if candidates.empty?
    exact = candidates.select { |c| c[:exact] }
    pool = exact.any? ? exact : candidates
    best = pool.min_by { |c| [c[:diff], c[:prefer_1220]] }
    { material_id: best[:material_id], thickness_mm: best[:thickness_mm] }
  end

  def self.sheet_unit?(size_str)
    return false unless size_str
    size_str.to_s.match?(/^\d+\s*[xX×]\s*\d+$/)
  end

  def self.linear_unit?(size_str)
    return false unless size_str
    size_str.to_s.match?(/^\d+\s*[xX×]\s*\d+\s*[xX×]\s*\d+$/)
  end

  def self.load_csv_data(path)
    return {} unless path && File.exist?(path)
    data = {}
    CSV.foreach(path, headers: true, encoding: "UTF-8") do |row|
      id = row["id"]&.to_s&.strip
      next if id.nil? || id.empty?
      data[id] = {
        name: (row["名前"] || row["name"] || "").to_s.strip,
        price: (row["値段"] || row["price"] || 0).to_f,
        memo: (row["メモ"] || row["memo"] || "").to_s.strip,
        material_id: (row["material_id"] || "").to_s.strip,
        finish_id: (row["finish_id"] || "").to_s.strip,
        material_name: (row["material_name"] || "").to_s.strip,
        thickness_mm: (row["thickness_mm"] || "").to_s.strip
      }
    end
    data
  rescue StandardError
    {}
  end

  # マージ: エンティティ属性 > CSV > 既定値。parent_entity は DC の panel_thickness 参照用
  def self.merge_for_update(old_item, new_item, csv_data, entity = nil, parent_entity = nil)
    id = new_item[:id].to_s
    csv_row = csv_data&.dig(id)

    # エンティティ属性を優先
    if entity && entity.respond_to?(:get_attribute)
      new_item[:thickness_mm] = entity.get_attribute(DICT, ATTR_THICKNESS_MM, "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || (csv_row&.dig(:thickness_mm) || old_item&.dig(:thickness_mm) || "")
      new_item[:material_id] = entity.get_attribute(DICT, ATTR_MATERIAL_ID, "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || (csv_row&.dig(:material_id) || old_item&.dig(:material_id) || "")
      new_item[:finish_id] = entity.get_attribute(DICT, ATTR_FINISH_ID, "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || (csv_row&.dig(:finish_id) || old_item&.dig(:finish_id) || "")
      new_item[:material_name] = entity.get_attribute(DICT, ATTR_MATERIAL_NAME, "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || (csv_row&.dig(:material_name) || old_item&.dig(:material_name) || "")
      new_item[:name] = entity.get_attribute(DICT, ATTR_NAME, "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || (csv_row ? csv_row[:name].to_s.strip : (old_item&.dig(:name).to_s || ""))
      price_attr = entity.get_attribute(DICT, ATTR_PRICE, -999)
      new_item[:price] = ((price_attr != -999 && price_attr >= 0 ? price_attr.to_f : nil) || (csv_row ? csv_row[:price].to_f : (old_item&.dig(:price).to_f || 0))).to_f.round
      new_item[:memo] = entity.get_attribute(DICT, ATTR_MEMO, "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || (csv_row ? csv_row[:memo].to_s.strip : (old_item&.dig(:memo).to_s || ""))
      hidden_val = entity.get_attribute(DICT, ATTR_HIDDEN, false)
      new_item[:hidden] = hidden_val == true || hidden_val.to_s.downcase == "true" || hidden_val == 1
      new_item[:price_override] = entity.get_attribute(DICT, ATTR_PRICE_OVERRIDE, false) == true
      new_item[:last_calc_bounds] = entity.get_attribute(DICT, ATTR_LAST_CALC_BOUNDS, "").to_s.strip
      new_item[:last_calc_bounds] = nil if new_item[:last_calc_bounds].empty?
      # 形状: ユーザーが属性で上書きしていればそれを優先（入れ子は構造で固定のため除外）
      unless new_item[:children]&.any?
        attr_shape = entity.get_attribute(DICT, ATTR_SHAPE_TYPE, "").to_s.strip
        if %w[beam board freeform other].include?(attr_shape)
          new_item[:shape_type] = attr_shape
        end
      end
    else
      new_item[:thickness_mm] = csv_row ? csv_row[:thickness_mm].to_s.strip : (old_item&.dig(:thickness_mm).to_s || "")
      new_item[:material_id] = csv_row ? csv_row[:material_id].to_s.strip : (old_item&.dig(:material_id).to_s || "")
      new_item[:finish_id] = csv_row ? csv_row[:finish_id].to_s.strip : (old_item&.dig(:finish_id).to_s || "")
      new_item[:material_name] = csv_row ? csv_row[:material_name].to_s.strip : (old_item&.dig(:material_name).to_s || "")
      new_item[:name] = csv_row ? csv_row[:name].to_s.strip : (old_item&.dig(:name).to_s || "")
      new_item[:price] = (csv_row ? csv_row[:price].to_f : (old_item&.dig(:price).to_f || 0)).round
      new_item[:memo] = csv_row ? csv_row[:memo].to_s.strip : (old_item&.dig(:memo).to_s || "")
      new_item[:hidden] = old_item&.dig(:hidden) ? true : false
      new_item[:price_override] = false
      new_item[:last_calc_bounds] = nil
    end

    # 棒材で下地材が未選択の場合、断面サイズから自動マッチング
    if new_item[:shape_type] == "beam" && entity && new_item[:material_id].to_s.strip.empty?
      match = find_best_beam_match_by_section(entity)
      new_item[:material_id] = match[:material_id] if match
    end

    # 板材で厚み・下地材が未選択の場合、モデルから自動抽出・選択（DC の panel_thickness を優先）
    if new_item[:shape_type] == "board" && entity
      if new_item[:thickness_mm].to_s.strip.empty?
        model_mm = board_thickness_mm(entity, parent_entity)
        if model_mm && model_mm > 0
          # 厚みを自動設定（選択肢に近いものを探す）
          opts = board_thickness_options
          nearest = opts.min_by { |t| (t - model_mm).abs } if opts.any?
          new_item[:thickness_mm] = nearest.to_s if nearest
        end
      end
      if new_item[:material_id].to_s.strip.empty?
        match = find_best_board_match_by_thickness(entity, parent_entity)
        if match
          new_item[:material_id] = match[:material_id]
          new_item[:thickness_mm] = match[:thickness_mm].to_s if new_item[:thickness_mm].to_s.strip.empty?
        end
      end
    end

    # 価格計算（棒・面状）。bounds 変化時も再計算（DC の板厚変更など）
    bounds_changed = false
    if entity && new_item[:last_calc_bounds].to_s.strip != "" && !new_item[:price_override]
      current_sig = bounds_signature(entity)
      bounds_changed = (current_sig && new_item[:last_calc_bounds] != current_sig)
    end
    need_recalc = new_item[:price].to_f <= 0 || bounds_changed

    if %w[beam board].include?(new_item[:shape_type]) && need_recalc && entity
      # bounds 変化時は板厚・素材を再マッチ（DC の panel_thickness 変更対応）
      if new_item[:shape_type] == "board" && bounds_changed
        model_mm = board_thickness_mm(entity, parent_entity)
        if model_mm && model_mm > 0
          opts = board_thickness_options
          nearest = opts.min_by { |t| (t - model_mm).abs } if opts.any?
          new_item[:thickness_mm] = nearest.to_s if nearest
        end
        match = find_best_board_match_by_thickness(entity, parent_entity)
        if match
          new_item[:material_id] = match[:material_id]
          new_item[:thickness_mm] = match[:thickness_mm].to_s if new_item[:thickness_mm].to_s.strip.empty?
        end
      end
      calc = compute_price_for_item(new_item, entity)
      if calc && calc > 0
        new_item[:price] = calc.round
        if entity.respond_to?(:set_attribute)
          entity.set_attribute(DICT, ATTR_LAST_CALC_BOUNDS, bounds_signature(entity).to_s)
          entity.set_attribute(DICT, ATTR_PRICE_OVERRIDE, false)
        end
      end
    end

    return unless new_item[:children]&.any?
    new_item[:children].each do |new_c|
      old_c = (old_item&.dig(:children) || []).find { |o| o[:id].to_s == new_c[:id].to_s }
      child_entity = entity && find_entity_by_pid(new_c[:id])
      merge_for_update(old_c, new_c, csv_data, child_entity, entity)
    end
  end

  # リストの内容を CSV に書き込む（画面上の情報を CSV に反映）
  def self.update_list(edits_json = nil)
    save_csv(edits_json)
  end

  def self.update_single_item(id)
    return unless model
    old_item = find_item_in_tree(list_items, id.to_s)
    return unless old_item
    entity = find_entity_by_pid(id)
    unless entity
      UI.messagebox("モデル内でオブジェクトが見つかりません")
      return
    end
    new_item = build_list_item_tree(entity)
    merge_for_update(old_item, new_item, nil, entity)
    replace_item_in_tree(list_items, id.to_s, new_item)
    refresh_panel
  end

  def self.replace_item_in_tree(items, target_id, new_item)
    items.each_with_index do |it, i|
      if it[:id].to_s == target_id
        items[i] = new_item
        return true
      end
      next unless it[:children]&.any?
      return true if replace_item_in_tree(it[:children], target_id, new_item)
    end
    false
  end

  def self.update_shape_type(id, shape_type)
    return unless model
    item = find_item_in_tree(list_items, id.to_s)
    return unless item
    return if item[:children]&.any? # 入れ子は変更不可
    valid = %w[beam board freeform other]
    return unless valid.include?(shape_type.to_s.strip)
    item[:shape_type] = shape_type.to_s.strip
    entity = find_entity_by_pid(id)
    entity&.set_attribute(DICT, ATTR_SHAPE_TYPE, shape_type.to_s.strip) if entity&.respond_to?(:set_attribute)
    propagate_nested_prices!(list_items)
    refresh_panel
  end

  def self.toggle_item_visibility(id)
    return unless model
    item = find_item_in_tree(list_items, id.to_s)
    return unless item
    entity = find_entity_by_pid(id)
    new_val = !item[:hidden]
    item[:hidden] = new_val
    if entity && entity.respond_to?(:set_attribute)
      entity.set_attribute(DICT, ATTR_HIDDEN, new_val)
    end
    refresh_panel
  end

  def self.find_entity_by_pid(pid)
    return nil unless model
    result = model.find_entity_by_persistent_id(pid.to_i)
    result.is_a?(Array) ? result.first : result
  rescue StandardError
    nil
  end

  def self.zoom_to_item(payload)
    id = nil
    selected_ids = nil
    if payload.is_a?(String) && payload.to_s.strip.start_with?("{")
      h = JSON.parse(payload) rescue {}
      id = (h["id"] || h[:id]).to_s
      selected_ids = h["selectedIds"] || h[:selectedIds]
    end
    id = payload.to_s if id.nil? || id.empty?
    return if id.to_s.strip.empty?

    entity = find_entity_by_pid(id)
    unless entity
      show_panel_status("該当オブジェクトが見つかりません。同期ボタンでリストを再取得してください。")
      return
    end
    @selection_sync_from_list = true
    model.selection.clear
    model.selection.add(entity)
    zoom_to_entity_with_origin_target(entity)
    model.active_view.refresh if model.active_view
    # パネルのハイライト・選択状態を同期（選択解除時は空のまま）
    list_ids = all_list_ids.to_set
    ids_to_highlight = (selected_ids || [id]).map(&:to_s).select { |s| list_ids.include?(s) }.uniq
    ids_to_highlight = [id.to_s] if ids_to_highlight.empty? && (selected_ids.nil? || !selected_ids.empty?)
    expanded = ancestor_ids_for_ids(ids_to_highlight)
    refresh_panel(highlight_ids: ids_to_highlight, selected_ids: ids_to_highlight, expanded_ids: expanded)
  ensure
    @selection_sync_from_list = false
  end

  # 入れ子グループ含め、エンティティのモデル座標系での変換を取得
  def self.transformation_to_model_space(entity)
    tr = entity.transformation
    ent = entity
    loop do
      container = ent.parent
      break unless container
      owner = container.respond_to?(:parent) ? container.parent : nil
      break unless owner && owner.respond_to?(:transformation)
      break if owner.is_a?(Sketchup::Model)
      tr = owner.transformation * tr
      ent = owner
    end
    tr
  end

  # transformation.origin をカメラターゲットにしてエンティティをフレーム内に収める
  def self.zoom_to_entity_with_origin_target(entity)
    view = model.active_view
    cam = view.camera
    # 入れ子の場合はモデル座標系に変換
    tr_to_model = transformation_to_model_space(entity)
    target = tr_to_model.origin
    bb = entity.bounds
    # bounds は親座標系なので、親→モデルの変換を適用（tr_to_model = entity_tr * parent_to_model）
    tr_parent_to_model = entity.transformation.inverse * tr_to_model
    up = cam.up
    # カメラ方向を維持しつつ、ターゲットを transformation.origin に変更
    dir = (cam.eye - cam.target)
    dir = Geom::Vector3d.new(0, 0, 1) if dir.length < 1.0e-6
    dir.normalize! if dir.length > 1.0e-6
    # ターゲットから最も遠い bounds の角までの距離（モデル座標で）
    max_dist = 0.0
    8.times do |i|
      corner_in_model = tr_parent_to_model * bb.corner(i)
      max_dist = [max_dist, corner_in_model.distance(target)].max
    end
    max_dist = 1.0 if max_dist < 1.0e-6
    # FOV に応じた距離（オブジェクトが画面に収まるように）
    fov_rad = (cam.fov || 30) * Math::PI / 180.0
    distance = max_dist / Math.tan(fov_rad * 0.5) * 1.2
    distance = [distance, max_dist * 2].max
    eye = target.offset(dir, distance)
    new_cam = Sketchup::Camera.new(eye, target, up, cam.perspective?, cam.fov || 30)
    view.camera = new_cam
  rescue StandardError => e
    # フォールバック: 従来の選択ズーム
    begin
      Sketchup.send_action("viewZoomToSelection:")
    rescue StandardError
      view.zoom(model.selection)
    end
  end

  # リストに含まれる全IDを収集（入れ子含む）
  def self.all_list_ids
    result = []
    list_items.each { |it| collect_ids_from_item(it, result) }
    result.uniq
  end

  def self.collect_ids_from_item(item, result)
    id = item[:id].to_s
    result << id unless id.empty?
    (item[:children] || []).each { |c| collect_ids_from_item(c, result) }
  end

  # 指定IDの祖先IDを収集（入れ子の子を表示するため展開が必要な親）
  def self.ancestor_ids_for_ids(target_ids, items = list_items, ancestors = [], result = [])
    target_set = target_ids.to_set
    items.each do |it|
      id = it[:id].to_s
      children = it[:children] || []
      if target_set.include?(id)
        result.concat(ancestors)
      end
      ancestor_ids_for_ids(target_ids, children, ancestors + [id], result) if children.any?
    end
    result.uniq
  end

  # 3D選択に応じてリストをハイライト
  def self.sync_list_from_selection
    return unless model && @dialog && @dialog.visible?
    return if @selection_sync_from_list
    list_ids_set = all_list_ids.to_set
    sel_ids = model.selection.select { |e| e.is_a?(Sketchup::Group) || e.is_a?(Sketchup::ComponentInstance) }
    matched = sel_ids.filter_map do |e|
      next nil unless e.respond_to?(:persistent_id)
      pid = e.persistent_id.to_s
      pid if list_ids_set.include?(pid)
    end.uniq
    expanded = ancestor_ids_for_ids(matched)
    refresh_panel(highlight_ids: matched, selected_ids: matched, expanded_ids: expanded)
  end

  def self.find_item_in_tree(items, target_id)
    items.each do |it|
      return it if it[:id].to_s == target_id.to_s
      found = find_item_in_tree(it[:children] || [], target_id)
      return found if found
    end
    nil
  end

  # ---------- 形状判定 ----------
  # exclude_hidden: true のとき Show_Legs 等で非表示の子をスキップ（価格集計に含めない）
  def self.collect_children(entity, exclude_hidden: false)
    entities = entity.is_a?(Sketchup::Group) ? entity.entities : entity.definition.entities
    children = entities.select { |e| e.is_a?(Sketchup::Group) || e.is_a?(Sketchup::ComponentInstance) }
    children = children.reject { |e| e.respond_to?(:hidden?) && e.hidden? } if exclude_hidden
    children
  end

  def self.board_like?(e)
    return false unless e.is_a?(Sketchup::Group) || e.is_a?(Sketchup::ComponentInstance)
    return false if collect_children(e).any?
    bb = e.bounds
    dims = [bb.width.to_mm, bb.height.to_mm, bb.depth.to_mm].sort
    thickness, dim1, dim2 = dims[0], dims[1], dims[2]
    return false if dim1 <= 0
    thickness >= 0 && thickness <= BOARD_THICKNESS_MAX_MM &&
      dim1 >= BOARD_MIN_DIM_MM && dim2 >= BOARD_MIN_DIM_MM &&
      thickness / dim1 < BOARD_THICKNESS_RATIO_MAX
  end

  def self.beam_like?(e)
    return false unless e.is_a?(Sketchup::Group) || e.is_a?(Sketchup::ComponentInstance)
    return false if collect_children(e).any?
    bb = e.bounds
    dims = [bb.width.to_mm, bb.height.to_mm, bb.depth.to_mm].sort
    min_d, mid_d, max_d = dims[0], dims[1], dims[2]
    return false if mid_d <= 0
    min_d >= BEAM_MIN_D_MM && max_d / mid_d >= BEAM_ASPECT_RATIO_MIN
  end

  def self.face_count(entity)
    return 0 unless entity.is_a?(Sketchup::Group) || entity.is_a?(Sketchup::ComponentInstance)
    entities = entity.is_a?(Sketchup::Group) ? entity.entities : entity.definition.entities
    count = entities.grep(Sketchup::Face).size
    entities.grep(Sketchup::Group).each { |g| count += face_count(g) }
    entities.grep(Sketchup::ComponentInstance).each { |c| count += face_count(c) }
    count
  end

  def self.beam_length_mm(entity)
    return 0 unless entity.respond_to?(:bounds)
    [entity.bounds.width.to_mm, entity.bounds.height.to_mm, entity.bounds.depth.to_mm].max
  end

  # 棒材エンティティの断面サイズ [幅mm, 高さmm]（長さ以外の2辺、昇順）
  def self.beam_section_mm(entity)
    return nil unless entity&.respond_to?(:bounds)
    dims = [entity.bounds.width.to_mm, entity.bounds.height.to_mm, entity.bounds.depth.to_mm].sort
    return nil if dims[0] <= 0 || dims[1] <= 0
    [dims[0], dims[1]]
  end

  # 棒材CSVのunit_size（30x40x2700）から断面 [幅, 高さ] を取得
  def self.parse_beam_section(row)
    return nil unless row
    sz = (row["unit_size"] || "").to_s.strip
    m = sz.match(/^(\d+\.?\d*)\s*[xX×]\s*(\d+\.?\d*)\s*[xX×]\s*(\d+\.?\d*)$/)
    return nil unless m
    [m[1].to_f, m[2].to_f].sort
  end

  # 断面が最も近い棒材の material_id を返す（同一断面優先）
  def self.find_best_beam_match_by_section(entity)
    target = beam_section_mm(entity)
    return nil unless target && target.size == 2 && target[0] > 0 && target[1] > 0
    candidates = raw_beam_options.map do |r|
      row_section = parse_beam_section(r)
      next nil unless row_section && row_section.size == 2
      diff_w = (row_section[0] - target[0]).abs
      diff_h = (row_section[1] - target[1]).abs
      diff = diff_w + diff_h
      exact = diff_w < 0.01 && diff_h < 0.01
      { material_id: r["material_id"].to_s.strip, diff: diff, exact: exact }
    end.compact
    return nil if candidates.empty?
    exact = candidates.select { |c| c[:exact] }
    best = exact.any? ? exact.min_by { |c| c[:diff] } : candidates.min_by { |c| c[:diff] }
    { material_id: best[:material_id] }
  end

  def self.board_area_mm2(entity)
    return 0 unless entity.respond_to?(:bounds)
    dims = [entity.bounds.width.to_mm, entity.bounds.height.to_mm, entity.bounds.depth.to_mm].sort
    return 0 if dims[1] <= 0 || dims[2] <= 0
    dims[1] * dims[2]
  end

  # 板材の縦横サイズキー（短辺x長辺 mm）
  def self.board_dims_key(entity)
    return nil unless entity&.respond_to?(:bounds)
    dims = [entity.bounds.width.to_mm, entity.bounds.height.to_mm, entity.bounds.depth.to_mm].sort
    return nil if dims[1] <= 0 || dims[2] <= 0
    "#{dims[1].round}×#{dims[2].round}"
  end

  # 選択IDのうち、他選択の子孫でないものだけ返す（二重カウント防止）
  def self.selection_roots(selected_ids)
    ids = selected_ids.uniq.map(&:to_s)
    ids.reject do |desc_id|
      ids.any? do |anc_id|
        next false if anc_id == desc_id
        anc = find_item_in_tree(list_items, anc_id)
        anc && anc[:children]&.any? && find_item_in_tree(anc[:children], desc_id)
      end
    end
  end

  def self.compute_beam_board_stats(selected_ids)
    result = { beam_count: 0, beam_total: 0.0, board_count: 0, board_area: 0.0, total_price: 0.0, breakdown: [] }
    beam_by_material = {}   # material_id => { total_mm, unit_len_mm }
    board_by_material = {}  # material_id => { total_mm2, sheet_m2 }
    board_by_size = {}      # "910×1820" => count
    finish_by_usage = {}    # finish_id => total_mm2
    item_details = []
    item_details_seen = {}
    breakdown_seen = {}
    root_ids = selection_roots(selected_ids)
    root_ids.each do |id|
      item = find_item_in_tree(list_items, id.to_s)
      next unless item
      next if item[:hidden]
      collect_item_details_recursive(item, item_details, item_details_seen)
      r = collect_beam_board_from_item(item)
      result[:beam_count] += r[:beam_count]
      result[:beam_total] += r[:beam_total]
      result[:board_count] += r[:board_count]
      result[:board_area] += r[:board_area]
      result[:total_price] += (item[:price] || 0).to_f
      collect_breakdown_recursive(item, result[:breakdown], breakdown_seen)
      collect_material_usage(item, beam_by_material, board_by_material, finish_by_usage)
      collect_board_sizes(item, board_by_size)
    end
    board_specs = compute_board_specs(board_by_material)
    beam_specs = compute_beam_specs(beam_by_material)
    finish_specs = compute_finish_specs(finish_by_usage)
    bounds = compute_selection_bounds_mm(selected_ids)
    {
      beam_count: result[:beam_count],
      beam_total_mm: result[:beam_total].round(1),
      board_count: result[:board_count],
      board_area_mm2: result[:board_area].round(1),
      board_sizes: board_by_size,
      bounds_mm: bounds,
      total_price: result[:total_price].round,
      breakdown: result[:breakdown],
      board_specs: board_specs,
      beam_specs: beam_specs,
      finish_specs: finish_specs,
      item_details: item_details
    }
  end

  # 選択項目全体の外接直方体サイズ（選択された各コンポーネントの bounds をモデル座標で合成）
  def self.compute_selection_bounds_mm(selected_ids)
    return nil if selected_ids.nil? || selected_ids.empty?
    corners = []
    root_ids = selection_roots(selected_ids)
    root_ids.each do |id|
      item = find_item_in_tree(list_items, id.to_s)
      next unless item && !item[:hidden]
      entity = find_entity_by_pid(item[:id])
      next unless entity&.respond_to?(:bounds)
      tr = transformation_to_model_space(entity)
      bb = entity.bounds
      8.times { |i| corners << (tr * bb.corner(i)) }
    end
    return nil if corners.empty?
    xs = corners.map { |c| c.x }
    ys = corners.map { |c| c.y }
    zs = corners.map { |c| c.z }
    to_mm = ->(v) { v.respond_to?(:to_mm) ? v.to_mm : (v * 25.4) }
    {
      width_mm: to_mm.call(xs.max - xs.min).round(1),
      height_mm: to_mm.call(zs.max - zs.min).round(1),  # Z軸 = 高さ
      depth_mm: to_mm.call(ys.max - ys.min).round(1)
    }
  end

  def self.collect_board_sizes(item, board_by_size)
    return if item.nil? || item[:hidden]
    if item[:children]&.any?
      item[:children].each { |c| collect_board_sizes(c, board_by_size) unless c.nil? || c[:hidden] }
      return
    end
    return unless item[:shape_type] == "board"
    entity = find_entity_by_pid(item[:id])
    return unless entity
    key = board_dims_key(entity)
    return unless key
    board_by_size[key] ||= 0
    board_by_size[key] += 1
  end

  def self.collect_material_usage(item, beam_by_material, board_by_material, finish_by_usage)
    return if item.nil?
    return if item[:hidden]
    # 入れ子: 全子孫を再帰的に走査（同じコンポーネントの複数インスタンスは別々にカウントするため seen は使わない）
    if item[:children]&.any?
      item[:children].each do |c|
        collect_material_usage(c, beam_by_material, board_by_material, finish_by_usage) unless c.nil? || c[:hidden]
      end
      return
    end
    entity = find_entity_by_pid(item[:id])
    return unless entity
    case item[:shape_type]
    when "beam"
      mid = item[:material_id].to_s.strip
      return if mid.empty?
      raw = raw_by_id[mid]
      return unless raw && linear_unit?(raw["unit_size"])
      len = raw["unit_size"].to_s.match(/\d+\s*[xX×]\s*\d+\s*[xX×]\s*(\d+)/)&.[](1)&.to_f || 1000
      len = 1000.0 if len <= 0
      beam_by_material[mid] ||= { total_mm: 0.0, unit_len_mm: len }
      beam_by_material[mid][:total_mm] += beam_length_mm(entity)
    when "board"
      mid = item[:material_id].to_s.strip
      raw = raw_by_id[mid]
      if raw && sheet_unit?(raw["unit_size"])
        m = raw["unit_size"].to_s.match(/(\d+)\s*[xX×]\s*(\d+)/)
        if m
          sheet_m2 = (m[1].to_f / 1000.0) * (m[2].to_f / 1000.0)
          board_by_material[mid] ||= { total_mm2: 0.0, sheet_m2: sheet_m2 }
          board_by_material[mid][:total_mm2] += board_area_mm2(entity)
        end
      end
      fid = item[:finish_id].to_s.strip
      if fid && !fid.empty?
        finish_by_usage[fid] ||= 0.0
        finish_by_usage[fid] += board_area_mm2(entity)
      end
    end
  end

  def self.compute_board_specs(board_by_material)
    board_by_material.map do |mid, h|
      total_m2 = (h[:total_mm2] || 0) / 1_000_000.0
      sheet_m2 = h[:sheet_m2] || 0.001
      sheets_needed = sheet_m2 > 0 ? (total_m2 / sheet_m2).ceil : 0
      used_m2 = sheets_needed * sheet_m2
      leftover_m2 = (used_m2 - total_m2).round(6)
      leftover_m2 = 0 if leftover_m2 < 0
      raw = raw_by_id[mid]
      name = raw ? (raw["name"] || mid).to_s.strip : mid.to_s
      { material_id: mid, material_name: name, total_m2: total_m2.round(4), sheet_m2: sheet_m2.round(4), sheets_needed: sheets_needed, leftover_m2: leftover_m2 }
    end
  end

  def self.compute_beam_specs(beam_by_material)
    beam_by_material.map do |mid, h|
      total_mm = (h[:total_mm] || 0).round(1)
      unit_len = h[:unit_len_mm] || 1000
      beams_needed = unit_len > 0 ? (total_mm / unit_len).ceil : 0
      used_mm = beams_needed * unit_len
      leftover_mm = (used_mm - total_mm).round(1)
      leftover_mm = 0 if leftover_mm < 0
      raw = raw_by_id[mid]
      name = raw ? (raw["name"] || mid).to_s.strip : mid.to_s
      { material_id: mid, material_name: name, total_mm: total_mm, unit_len_mm: unit_len.round, beams_needed: beams_needed, leftover_mm: leftover_mm }
    end
  end

  def self.compute_finish_specs(finish_by_usage)
    finish_by_usage.map do |fid, total_mm2|
      total_m2 = total_mm2 / 1_000_000.0
      fin = finish_by_id[fid]
      coverage = fin && fin["m2_coverage_per_unit"].to_f > 0 ? fin["m2_coverage_per_unit"].to_f : 1.0
      units_needed = (total_m2 / coverage).ceil
      used_m2 = units_needed * coverage
      leftover_m2 = (used_m2 - total_m2).round(6)
      leftover_m2 = 0 if leftover_m2 < 0
      name = fin ? (fin["name"] || fid).to_s.strip : fid.to_s
      { finish_id: fid, finish_name: name, total_m2: total_m2.round(4), coverage_per_unit: coverage, units_needed: units_needed, leftover_m2: leftover_m2 }
    end
  end

  def self.build_calc_detail(item)
    name = (item[:name] || "").to_s.strip
    name = "（無名）" if name.empty?
    price = (item[:price] || 0).to_f.round
    st = item[:shape_type].to_s
    if st == "nested" && item[:children]&.any?
      { name: name, price: price, detail: "子項目の合計" }
    elsif st == "beam"
      entity = find_entity_by_pid(item[:id])
      raw = entity ? raw_by_id[item[:material_id].to_s.strip] : nil
      if raw && linear_unit?(raw["unit_size"])
        len_mm = entity ? beam_length_mm(entity) : 0
        unit_len = raw["unit_size"].to_s.match(/\d+\s*[xX×]\s*\d+\s*[xX×]\s*(\d+)/)&.[](1)&.to_f || 1
        unit_len = 1000.0 if unit_len <= 0
        price_per_m = (raw["unit_price"].to_f / (unit_len / 1000.0)) * (1 + raw["loss_rate"].to_f / 100.0)
        { name: name, price: price, detail: "長さ #{len_mm.round} mm × ¥#{price_per_m.round}/m = ¥#{price}" }
      else
        { name: name, price: price, detail: "直接入力" }
      end
    elsif st == "board"
      entity = find_entity_by_pid(item[:id])
      raw = entity ? raw_by_id[item[:material_id].to_s.strip] : nil
      if raw && sheet_unit?(raw["unit_size"])
        area_mm2 = entity ? board_area_mm2(entity) : 0
        m = raw["unit_size"].to_s.match(/(\d+)\s*[xX×]\s*(\d+)/)
        if m
          sheet_area = (m[1].to_f / 1000.0) * (m[2].to_f / 1000.0)
          price_per_m2 = sheet_area > 0 ? (raw["unit_price"].to_f / sheet_area) * (1 + raw["loss_rate"].to_f / 100.0) : 0
          fin = finish_by_id[item[:finish_id].to_s.strip]
          price_per_m2 += fin["1m2_price"].to_f if fin
          { name: name, price: price, detail: "面積 #{(area_mm2 / 1_000_000.0).round(4)} m² × ¥#{price_per_m2.round}/m² = ¥#{price}" }
        else
          { name: name, price: price, detail: "直接入力" }
        end
      else
        { name: name, price: price, detail: "直接入力" }
      end
    else
      { name: name, price: price, detail: "直接入力" }
    end
  end

  # 入れ子選択時も子孫のサイズ・計算根拠を集計に含める
  def self.collect_item_details_recursive(item, out, seen = {})
    return if item.nil? || item[:hidden]
    key = item[:id].to_s
    return if seen[key]
    seen[key] = true
    out << { name: (item[:name] || "").to_s.strip, detail: (item[:detail] || "").to_s }
    (item[:children] || []).each { |c| collect_item_details_recursive(c, out, seen) }
  end

  def self.collect_breakdown_recursive(item, out, seen = {})
    return if item.nil? || item[:hidden]
    key = item[:id].to_s
    return if seen[key]
    seen[key] = true
    out << build_calc_detail(item)
    (item[:children] || []).each { |c| collect_breakdown_recursive(c, out, seen) }
  end

  def self.collect_beam_board_from_item(item, _seen = nil)
    return { beam_count: 0, beam_total: 0.0, board_count: 0, board_area: 0.0 } if item.nil?
    return { beam_count: 0, beam_total: 0.0, board_count: 0, board_area: 0.0 } if item[:hidden]
    # 入れ子: 全子孫を再帰的に集計（同じコンポーネントの複数インスタンスは別々にカウントするため seen は使わない）
    if item[:children]&.any?
      beam_c, beam_t, board_c, board_a = 0, 0.0, 0, 0.0
      item[:children].each do |c|
        next if c.nil? || c[:hidden]
        r = collect_beam_board_from_item(c)
        beam_c += r[:beam_count]; beam_t += r[:beam_total]; board_c += r[:board_count]; board_a += r[:board_area]
      end
      return { beam_count: beam_c, beam_total: beam_t, board_count: board_c, board_area: board_a }
    end
    entity = find_entity_by_pid(item[:id])
    return { beam_count: 0, beam_total: 0.0, board_count: 0, board_area: 0.0 } unless entity
    case item[:shape_type]
    when "beam" then { beam_count: 1, beam_total: beam_length_mm(entity), board_count: 0, board_area: 0.0 }
    when "board" then { beam_count: 0, beam_total: 0.0, board_count: 1, board_area: board_area_mm2(entity) }
    else { beam_count: 0, beam_total: 0.0, board_count: 0, board_area: 0.0 }
    end
  end

  # 形状判定: 入れ子状 → 棒 → 面状 → 自由形状(6面以上) → その他
  def self.determine_shape_type(e)
    return "nested" if collect_children(e).any?
    return "beam" if beam_like?(e)
    return "board" if board_like?(e)
    return "freeform" if face_count(e) >= FREEFORM_MIN_FACE_COUNT
    "other"
  end

  def self.build_list_item_tree(e)
    item = build_list_item(e)
    children = collect_children(e, exclude_hidden: true).map { |child| build_list_item_tree(child) }
    item[:children] = children if children.any?
    if children.empty?
      item[:shape_type] = determine_shape_type(e)
    else
      item[:shape_type] = "nested"
    end
    item
  end

  def self.build_list_item(e)
    pid = e.respond_to?(:persistent_id) ? e.persistent_id.to_s : object_id.to_s
    detail = (e.respond_to?(:bounds) ? size_string(e) : nil) || "—"
    {
      id: pid,
      type: entity_type_label(e),
      tag: entity_tag(e),
      name: "",
      detail: detail,
      price: 0,
      memo: "",
      material_id: "",
      finish_id: "",
      material_name: ""
    }
  end

  # 価格計算（棒・面状）
  def self.compute_price_for_item(item, entity)
    return nil unless entity
    case item[:shape_type]
    when "beam"
      raw = raw_by_id[item[:material_id].to_s.strip]
      return nil unless raw && linear_unit?(raw["unit_size"])
      len_m = beam_length_mm(entity) / 1000.0
      unit_len = raw["unit_size"].to_s.match(/\d+\s*[xX×]\s*\d+\s*[xX×]\s*(\d+)/)&.[](1)&.to_f || 1
      unit_len = 1000.0 if unit_len <= 0
      price_per_m = (raw["unit_price"].to_f / (unit_len / 1000.0)) * (1 + raw["loss_rate"].to_f / 100.0)
      price_per_m * len_m
    when "board"
      raw = raw_by_id[item[:material_id].to_s.strip]
      return nil unless raw && sheet_unit?(raw["unit_size"])
      area_m2 = board_area_mm2(entity) / 1_000_000.0
      m = raw["unit_size"].to_s.match(/(\d+)\s*[xX×]\s*(\d+)/)
      return nil unless m
      sheet_area = (m[1].to_f / 1000.0) * (m[2].to_f / 1000.0)
      return nil if sheet_area <= 0
      price_per_m2 = (raw["unit_price"].to_f / sheet_area) * (1 + raw["loss_rate"].to_f / 100.0)
      total = price_per_m2 * area_m2
      fin = finish_by_id[item[:finish_id].to_s.strip]
      total += (fin["1m2_price"].to_f * area_m2) if fin
      total
    else
      nil
    end
  end

  def self.format_error(e, context = "")
    lines = ["[#{context}] #{e.class}: #{e.message}"]
    (e.backtrace || []).first(8).each { |line| lines << "  #{line}" }
    lines.join("\n")
  end

  # ---------- オンラインアップデート ----------
  # バージョン文字列を比較（1.1.1 < 1.1.2 など）
  def self.version_newer?(remote, local)
    return false if remote.to_s.strip.empty?
    r_parts = remote.to_s.strip.split(/[.\-]/).map { |s| s.to_i }
    l_parts = local.to_s.strip.split(/[.\-]/).map { |s| s.to_i }
    max_len = [r_parts.size, l_parts.size].max
    max_len.times do |i|
      r = r_parts[i] || 0
      l = l_parts[i] || 0
      return true if r > l
      return false if r < l
    end
    false
  end

  def self.last_update_check_time
    Sketchup.read_default(DICT, UPDATE_CHECK_PREF_KEY, 0).to_i
  end

  def self.save_last_update_check_time
    Sketchup.write_default(DICT, UPDATE_CHECK_PREF_KEY, Time.now.to_i)
  end

  def self.should_check_update?
    return false if UPDATE_CHECK_URL.to_s.strip.empty? || UPDATE_CHECK_URL.include?("example.com")
    (Time.now.to_i - last_update_check_time) >= UPDATE_CHECK_INTERVAL
  end

  # ファイルをダウンロードして Plugins にインストール（自動更新）
  def self.download_and_install_update(files, remote_version)
    plugins_dir = File.dirname(__FILE__)
    base_url = (UPDATE_RAW_BASE || "").to_s.strip
    if base_url.empty?
      UI.messagebox("自動更新の設定がありません。ダウンロードページから手動で更新してください。")
      return
    end
    base_url += "/" unless base_url.end_with?("/")

    index_ref = [0]
    total = files.size

    download_next = lambda do
      if index_ref[0] >= total
        UI.start_timer(0, false) { UI.messagebox("M List を #{remote_version} に更新しました。\n\nSketchUp を再起動してください。") }
        return
      end
      file = files[index_ref[0]].to_s.strip
      next if file.empty?
      url = base_url + file.gsub(" ", "%20")
      dest = File.join(plugins_dir, file)
      req = Sketchup::Http::Request.new(url, Sketchup::Http::GET)
      req.start do |_request, response|
        if response && response.respond_to?(:status_code) && response.status_code == 200
          body = response.body
          if body && !body.to_s.empty?
            FileUtils.mkdir_p(File.dirname(dest))
            File.open(dest, "wb") { |f| f.write(body.respond_to?(:to_str) ? body.to_str : body.to_s) }
          end
        else
          UI.start_timer(0, false) { UI.messagebox("ダウンロード失敗: #{file}\n\nダウンロードページから手動で更新してください。") }
          return
        end
        index_ref[0] += 1
        download_next.call
      end
    rescue StandardError => e
      UI.start_timer(0, false) { UI.messagebox("更新エラー: #{e.message}\n\nダウンロードページから手動で更新してください。") }
    end

    download_next.call
  end

  # オンラインでバージョンを確認し、新しいバージョンがあれば通知
  def self.check_for_updates(manual: false)
    return if UPDATE_CHECK_URL.to_s.strip.empty? || UPDATE_CHECK_URL.include?("example.com")
    return unless manual || should_check_update?

    req = Sketchup::Http::Request.new(UPDATE_CHECK_URL, Sketchup::Http::GET)
    req.start do |_request, response|
      next unless response
      next unless response.respond_to?(:status_code) && response.status_code == 200

      body = response.body.to_s.strip
      next if body.empty?

      remote_version = nil
      download_url = UPDATE_DOWNLOAD_URL
      release_notes = ""

      begin
        data = JSON.parse(body)
        remote_version = (data["version"] || data[:version]).to_s.strip
        download_url = (data["download_url"] || data[:download_url]).to_s.strip
        download_url = UPDATE_DOWNLOAD_URL if download_url.empty?
        release_notes = (data["release_notes"] || data[:release_notes]).to_s.strip
      rescue JSON::ParserError
        remote_version = body.lines.first.to_s.strip
      end

      next if remote_version.empty?
      save_last_update_check_time

      if version_newer?(remote_version, VERSION)
        msg = "M List の新しいバージョン（#{remote_version}）が利用可能です。\n\n"
        msg += "現在のバージョン: #{VERSION}\n\n"
        msg += "#{release_notes}\n\n" unless release_notes.empty?
        msg += "今すぐ自動更新しますか？\n（完了後、SketchUpを再起動してください）\n\n"
        msg += "[はい] 自動更新\n[いいえ] ダウンロードページを開く"
        result = UI.messagebox(msg, MB_YESNO)
        if result == 6 # IDYES
          files = (data["files"] || data[:files] || [] rescue [])
          if files.is_a?(Array) && files.any?
            download_and_install_update(files, remote_version)
          else
            UI.openURL(download_url)
          end
        else
          UI.openURL(download_url)
        end
      elsif manual
        UI.messagebox("M List は最新バージョン（#{VERSION}）です。")
      end
    end
  rescue StandardError => e
    UI.messagebox("アップデート確認エラー\n#{e.message}") if manual
  end

  # 入れ子状アイテムの価格を子の合計に伝播
  def self.propagate_nested_prices!(items)
    items.each do |it|
      propagate_nested_prices!(it[:children] || []) if it[:children]&.any?
      if it[:children]&.any?
        it[:price] = it[:children].sum { |c| (c[:price] || 0).to_f }.round
      end
    end
  end

  # 検証警告を収集（reassign_from_shape と共有）
  def self.collect_validation_warnings(items = list_items, path_prefix = "")
    warnings = []
    collect = lambda do |its, prefix|
      its.each do |it|
        next if it[:hidden]
        name = (it[:name] || "").to_s.strip
        name = "（無名）" if name.empty?
        path = prefix.empty? ? name : "#{prefix} > #{name}"

        if it[:children]&.any?
          child_sum = it[:children].sum { |c| (c[:price] || 0).to_f }.round
          parent_price = (it[:price] || 0).to_f.round
          warnings << { type: "parent_mismatch", path: path, msg: "親の値段(¥#{parent_price})が子の合計(¥#{child_sum})と一致していません" } if parent_price != child_sum
          collect.call(it[:children], path)
          next
        end

        entity = find_entity_by_pid(it[:id])
        unless entity
          warnings << { type: "entity_not_found", path: path, msg: "モデル内でオブジェクトが見つかりません（削除された可能性）" }
          next
        end

        case it[:shape_type].to_s
        when "beam"
          mid = it[:material_id].to_s.strip
          if mid.empty?
            warnings << { type: "beam_no_material", path: path, msg: "棒材ですが材料が未選択です" }
          else
            raw = raw_by_id[mid]
            unless raw && linear_unit?(raw["unit_size"])
              warnings << { type: "beam_material_invalid", path: path, msg: "材料「#{mid}」が素材CSVに存在しないか、規格が不正です" }
            else
              calc = compute_price_for_item(it, entity)
              if calc && calc > 0
                diff = ((it[:price] || 0).to_f - calc).abs
                warnings << { type: "beam_price_mismatch", path: path, msg: "値段(¥#{(it[:price] || 0).to_f.round})が計算値(¥#{calc.round})と異なります" } if diff > 0.5
              end
            end
          end
        when "board"
          mid = it[:material_id].to_s.strip
          fid = it[:finish_id].to_s.strip
          fin = fid.empty? ? nil : finish_by_id[fid]
          if fid && !fid.empty? && !fin
            warnings << { type: "board_finish_invalid", path: path, msg: "仕上材「#{fid}」が素材CSVに存在しません" }
          elsif fid && !fid.empty? && fin && fin["1m2_price"].to_f > 0 && mid.empty?
            warnings << { type: "board_finish_no_base", path: path, msg: "仕上材が選択されていますが、下地材が未設定のため価格に反映されていません" }
          end
          if mid && !mid.empty?
            raw = raw_by_id[mid]
            unless raw && sheet_unit?(raw["unit_size"])
              warnings << { type: "board_material_invalid", path: path, msg: "下地材「#{mid}」が素材CSVに存在しないか、規格が不正です" }
            else
              calc = compute_price_for_item(it, entity)
              if calc && calc > 0
                diff = ((it[:price] || 0).to_f - calc).abs
                warnings << { type: "board_price_mismatch", path: path, msg: "値段(¥#{(it[:price] || 0).to_f.round})が計算値(¥#{calc.round})と異なります（仕上材の反映漏れの可能性）" } if diff > 0.5
              end
            end
          end
        end
      end
    end
    collect.call(items, path_prefix)
    warnings
  end

  # 形状から再当て込み: 形状を再判定し、価格を再計算。include_overrides=false のときはユーザー上書き項目をスキップ
  def self.reassign_from_shape(include_overrides: true)
    return { fixed_count: 0, warnings: [] } if list_items.empty? || !model
    fixed_count = 0
    warnings = []

    model.start_operation("形状から再当て込み", true)
    begin
      reassign_leaf = lambda do |items, parent_entity = nil|
        items.each do |it|
          if it[:children]&.any?
            parent_ent = find_entity_by_pid(it[:id])
            reassign_leaf.call(it[:children], parent_ent)
            next
          end
          next if it[:hidden]
          entity = find_entity_by_pid(it[:id])
          next unless entity

          # 上書き項目をスキップ（include_overrides=false の場合）
          unless include_overrides
            next if entity.get_attribute(DICT, ATTR_PRICE_OVERRIDE, false) == true
          end

          # 形状を再判定（入れ子は除く）
          new_shape = determine_shape_type(entity)
          old_shape = it[:shape_type].to_s

          # 形状タイプが変わった場合: material_id / thickness_mm / finish_id をクリア
          if new_shape != old_shape
            it[:material_id] = ""
            it[:thickness_mm] = ""
            it[:finish_id] = ""
            it[:material_name] = ""
            entity.set_attribute(DICT, ATTR_MATERIAL_ID, "")
            entity.set_attribute(DICT, ATTR_THICKNESS_MM, "")
            entity.set_attribute(DICT, ATTR_FINISH_ID, "")
            entity.set_attribute(DICT, ATTR_MATERIAL_NAME, "")
          end

          it[:shape_type] = new_shape
          entity.set_attribute(DICT, ATTR_SHAPE_TYPE, new_shape) if entity.respond_to?(:set_attribute)

          # beam/board の場合: 素材が未選択なら自動マッチ（DC の panel_thickness を優先）、価格を再計算
          if %w[beam board].include?(new_shape)
            if new_shape == "beam" && it[:material_id].to_s.strip.empty?
              match = find_best_beam_match_by_section(entity)
              it[:material_id] = match[:material_id] if match
            end
            if new_shape == "board"
              if it[:thickness_mm].to_s.strip.empty?
                model_mm = board_thickness_mm(entity, parent_entity)
                opts = board_thickness_options
                nearest = opts.min_by { |t| (t - model_mm).abs } if model_mm && opts.any?
                it[:thickness_mm] = nearest.to_s if nearest
              end
              if it[:material_id].to_s.strip.empty?
                match = find_best_board_match_by_thickness(entity, parent_entity)
                if match
                  it[:material_id] = match[:material_id]
                  it[:thickness_mm] = match[:thickness_mm].to_s if it[:thickness_mm].to_s.strip.empty?
                end
              end
            end

            calc = compute_price_for_item(it, entity)
            if calc && calc > 0
              new_price = calc.round
              it[:price] = new_price
              entity.set_attribute(DICT, ATTR_PRICE, new_price)
              sig = bounds_signature(entity)
              entity.set_attribute(DICT, ATTR_LAST_CALC_BOUNDS, sig.to_s) if sig
              entity.set_attribute(DICT, ATTR_PRICE_OVERRIDE, false)
              fixed_count += 1
            end
          end
        end
      end
      reassign_leaf.call(list_items)

      propagate_nested_prices!(list_items)

      persist_to_entities = lambda do |items|
        items.each do |it|
          persist_to_entities.call(it[:children] || []) if it[:children]&.any?
          entity = find_entity_by_pid(it[:id])
          next unless entity&.respond_to?(:set_attribute)
          entity.set_attribute(DICT, ATTR_PRICE, (it[:price] || 0).to_f.round)
          entity.set_attribute(DICT, ATTR_SHAPE_TYPE, it[:shape_type].to_s) if it[:shape_type]
          entity.set_attribute(DICT, ATTR_MATERIAL_ID, (it[:material_id] || "").to_s)
          entity.set_attribute(DICT, ATTR_THICKNESS_MM, (it[:thickness_mm] || "").to_s)
          entity.set_attribute(DICT, ATTR_FINISH_ID, (it[:finish_id] || "").to_s)
          entity.set_attribute(DICT, ATTR_MATERIAL_NAME, (it[:material_name] || "").to_s)
        end
      end
      persist_to_entities.call(list_items)

      warnings = collect_validation_warnings

      model.commit_operation
    rescue StandardError
      model.abort_operation
      raise
    end

    { fixed_count: fixed_count, warnings: warnings }
  end

  # ---------- パネル ----------
  def self.panel_html_path
    File.join(plugin_root, "panel.html")
  end

  def self.serialize_item_for_panel(it)
    h = {
      id: it[:id].to_s, type: it[:type], tag: (it[:tag] || "").to_s, name: (it[:name] || "").to_s,
      detail: it[:detail], price: (it[:price] || 0).to_f.round, memo: (it[:memo] || "").to_s,
      material_id: (it[:material_id] || "").to_s, finish_id: (it[:finish_id] || "").to_s,
      material_name: (it[:material_name] || "").to_s, thickness_mm: (it[:thickness_mm] || "").to_s
    }
    h[:shape_type] = it[:shape_type] if it[:shape_type]
    h[:hidden] = it[:hidden] ? true : false
    h[:children] = (it[:children] || []).map { |c| serialize_item_for_panel(c) } if it[:children]&.any?
    h
  end

  def self.update_item_in_tree(items, target_id, key, value)
    items.each do |it|
      if it[:id].to_s == target_id.to_s
        it[key] = value
        return true
      end
      return true if update_item_in_tree(it[:children] || [], target_id, key, value)
    end
    false
  end

  def self.apply_edits_from_payload(edits_json)
    return if edits_json.to_s.strip.empty?
    edits = JSON.parse(edits_json)
    return unless model
    model.start_operation("編集を反映", true)
    begin
      edits.each do |id_str, data|
        next unless id_str && !id_str.to_s.empty?
        h = data.is_a?(Hash) ? data : {}
        update_item_in_tree(list_items, id_str, :name, (h["name"] || h[:name] || "").to_s.strip)
        update_item_in_tree(list_items, id_str, :price, ((h["price"] || h[:price] || 0).to_f).round)
        update_item_in_tree(list_items, id_str, :memo, (h["memo"] || h[:memo] || "").to_s.strip)
        update_item_in_tree(list_items, id_str, :material_id, (h["material_id"] || h[:material_id] || "").to_s.strip)
        update_item_in_tree(list_items, id_str, :thickness_mm, (h["thickness_mm"] || h[:thickness_mm] || "").to_s.strip)
        update_item_in_tree(list_items, id_str, :finish_id, (h["finish_id"] || h[:finish_id] || "").to_s.strip)
        update_item_in_tree(list_items, id_str, :material_name, (h["material_name"] || h[:material_name] || "").to_s.strip)
        # エンティティに属性を保存
        entity = find_entity_by_pid(id_str)
        if entity && entity.respond_to?(:set_attribute)
          entity.set_attribute(DICT, ATTR_NAME, (h["name"] || "").to_s.strip)
          entity.set_attribute(DICT, ATTR_PRICE, ((h["price"] || 0).to_f).round)
          entity.set_attribute(DICT, ATTR_MEMO, (h["memo"] || "").to_s.strip)
          entity.set_attribute(DICT, ATTR_MATERIAL_ID, (h["material_id"] || "").to_s.strip)
          entity.set_attribute(DICT, ATTR_THICKNESS_MM, (h["thickness_mm"] || "").to_s.strip)
          entity.set_attribute(DICT, ATTR_FINISH_ID, (h["finish_id"] || "").to_s.strip)
          entity.set_attribute(DICT, ATTR_MATERIAL_NAME, (h["material_name"] || "").to_s.strip)
        end
      end
      propagate_nested_prices!(list_items)
      model.commit_operation
    rescue StandardError
      model.abort_operation
    end
  rescue StandardError
  end

  def self.write_csv_to_path(path, persist: true, edits_json: nil)
    apply_edits_from_payload(edits_json) if edits_json && !edits_json.to_s.strip.empty?
    propagate_nested_prices!(list_items)
    total_price = 0.0
    rows = flatten_items_for_csv(list_items, 0)
    rows = merge_edits_into_csv_rows(rows, edits_json) if edits_json && !edits_json.to_s.strip.empty?
    CSV.open(path, "w", encoding: "UTF-8") do |csv|
      csv << %w[id タグ 名前 種別 形状 サイズ material_id thickness_mm finish_id material_name 値段 メモ]
      rows.each do |r|
        csv << r
        total_price += r[10].to_f  # 値段は11列目
      end
      csv << ["", "", "", "", "", "", "", "", "", "合計", total_price.round, ""]
    end
    persist_csv_path(path) if persist
  end

  def self.compute_csv_total
    flatten_items_for_csv(list_items, 0).sum { |r| r[10].to_f }.round
  end

  def self.save_project(edits_json = nil)
    return unless model
    apply_edits_from_payload(edits_json)
    folder = project_folder_path
    base_name = nil
    unless folder && Dir.exist?(folder)
      default_base = "プロジェクト_#{Time.now.strftime('%Y%m%d')}"
      default_dir = File.expand_path("~/Documents")
      path = UI.savepanel("プロジェクトを保存", default_dir, "#{default_base}.skp")
      return if !path || path.empty?
      path += ".skp" unless path.downcase.end_with?(".skp")
      parent_dir = File.dirname(path)
      base_name = File.basename(path, ".*")
      base_name = "model" if base_name.empty?
      folder = File.join(parent_dir, base_name)
      FileUtils.mkdir_p(folder)
      self.project_folder_path = folder
    end
    base_name ||= model.get_attribute(DICT, PROJECT_BASE_NAME_KEY, "").to_s.strip
    base_name = File.basename(folder) if base_name.empty?
    base_name = "model" if base_name.empty?
    skp_path = File.join(folder, "#{base_name}.skp")
    csv_path = File.join(folder, "#{base_name}_estimate.csv")
    model.start_operation("プロジェクト保存", true)
    begin
      persist_list_ids
      model.save(skp_path)
      write_csv_to_path(csv_path, persist: true, edits_json: edits_json)
      write_materials_to_project_folder(folder)
      self.project_folder_path = folder
      model.set_attribute(DICT, PROJECT_BASE_NAME_KEY, base_name)
      show_panel_status("プロジェクトに保存しました: #{base_name}")
      msg = "プロジェクトに保存しました。\n\n#{skp_path}\n\nCSV: #{csv_path}"
      msg += "\n\n合計: ¥#{compute_csv_total}" if list_items.any?
      UI.messagebox(msg)
    ensure
      model.commit_operation
    end
  rescue StandardError => e
    model.abort_operation if model
    UI.messagebox("プロジェクト保存エラー\n#{e.message}")
  end

  def self.save_csv(edits_json = nil)
    return unless model
    apply_edits_from_payload(edits_json)
    if list_items.empty?
      UI.messagebox("リストが空です。\n保存する項目を追加してください。")
      return
    end
    path = resolved_csv_path
    unless path
      UI.messagebox("CSVの保存先が決まっていません。\n\n「プロジェクト保存」を先に行ってください。")
      return
    end
    write_csv_to_path(path, persist: true, edits_json: edits_json)
    show_panel_status("CSVに保存しました: #{File.basename(path)}")
    UI.messagebox("CSVに保存しました。\n#{path}\n\n合計: ¥#{compute_csv_total}")
  rescue StandardError => e
    UI.messagebox("CSV保存エラー\n#{format_error(e, 'save_csv')}")
  end

  # CSV の内容をリストに反映する（CSV → リスト）
  def self.csv_sync
    return unless model
    path = resolved_csv_path
    unless path && File.exist?(path)
      UI.messagebox("CSVファイルが見つかりません。\n\n「プロジェクト保存」または「保存」を先に行ってください。")
      return
    end
    csv_data = load_csv_data(path)
    return UI.messagebox("CSVに有効なデータがありません。") if csv_data.empty?
    count_ref = [0]
    apply_csv_to_tree(list_items, csv_data, count_ref)
    refresh_panel
    show_panel_status("CSVから #{count_ref[0]} 件をリストに反映しました。")
    UI.messagebox("CSVの内容をリストに反映しました。")
  rescue StandardError => e
    UI.messagebox("CSV同期エラー\n#{format_error(e, 'csv_sync')}")
  end

  # CSV をファイル選択で読み込み、リストに反映
  def self.load_csv_from_file
    return unless model
    if list_items.empty?
      UI.messagebox("リストが空です。\n先に「同期」ボタンでモデルからリストを取得してください。")
      return
    end
    path = UI.openpanel("CSVを読み込み", File.expand_path("~/Documents"), "*.csv")
    return if !path || path.empty?
    unless File.exist?(path)
      UI.messagebox("ファイルが見つかりません。")
      return
    end
    csv_data = load_csv_data(path)
    return UI.messagebox("CSVに有効なデータがありません。") if csv_data.empty?
    count_ref = [0]
    apply_csv_to_tree(list_items, csv_data, count_ref)
    refresh_panel
    show_panel_status("CSVから #{count_ref[0]} 件をリストに反映しました。")
    UI.messagebox("CSVの内容をリストに反映しました。")
  rescue StandardError => e
    UI.messagebox("CSV読み込みエラー\n#{format_error(e, 'load_csv_from_file')}")
  end

  # 別名保存（常に保存ダイアログを表示）
  def self.save_project_as(edits_json = nil)
    return unless model
    apply_edits_from_payload(edits_json)
    default_base = "プロジェクト_#{Time.now.strftime('%Y%m%d')}"
    default_dir = File.expand_path("~/Documents")
    path = UI.savepanel("プロジェクトを別名で保存", default_dir, "#{default_base}.skp")
    return if !path || path.empty?
    path += ".skp" unless path.downcase.end_with?(".skp")
    parent_dir = File.dirname(path)
    base_name = File.basename(path, ".*")
    base_name = "model" if base_name.empty?
    folder = File.join(parent_dir, base_name)
    FileUtils.mkdir_p(folder)
    skp_path = File.join(folder, "#{base_name}.skp")
    csv_path = File.join(folder, "#{base_name}_estimate.csv")
    model.start_operation("プロジェクト別名保存", true)
    begin
      persist_list_ids
      model.save(skp_path)
      write_csv_to_path(csv_path, persist: true, edits_json: edits_json)
      write_materials_to_project_folder(folder)
      self.project_folder_path = folder
      model.set_attribute(DICT, PROJECT_BASE_NAME_KEY, base_name)
      show_panel_status("プロジェクトに保存しました: #{base_name}")
      msg = "プロジェクトに保存しました。\n\n#{skp_path}\n\nCSV: #{csv_path}"
      msg += "\n\n合計: ¥#{compute_csv_total}" if list_items.any?
      UI.messagebox(msg)
    ensure
      model.commit_operation
    end
  rescue StandardError => e
    model.abort_operation if model
    UI.messagebox("プロジェクト保存エラー\n#{e.message}")
  end

  def self.apply_csv_to_tree(items, csv_data, count_ref)
    items.each do |it|
      id = it[:id].to_s
      row = csv_data[id]
      if row
        it[:name] = (row[:name] || "").to_s.strip
        it[:price] = ((row[:price] || 0).to_f).round
        it[:memo] = (row[:memo] || "").to_s.strip
        it[:material_id] = (row[:material_id] || "").to_s.strip
        it[:thickness_mm] = (row[:thickness_mm] || "").to_s.strip
        it[:finish_id] = (row[:finish_id] || "").to_s.strip
        it[:material_name] = (row[:material_name] || "").to_s.strip
        entity = find_entity_by_pid(id)
        if entity && entity.respond_to?(:set_attribute)
          entity.set_attribute(DICT, ATTR_NAME, it[:name])
          entity.set_attribute(DICT, ATTR_PRICE, it[:price].to_f.round)
          entity.set_attribute(DICT, ATTR_MEMO, it[:memo])
          entity.set_attribute(DICT, ATTR_MATERIAL_ID, it[:material_id])
          entity.set_attribute(DICT, ATTR_THICKNESS_MM, it[:thickness_mm])
          entity.set_attribute(DICT, ATTR_FINISH_ID, it[:finish_id])
          entity.set_attribute(DICT, ATTR_MATERIAL_NAME, it[:material_name])
          entity.set_attribute(DICT, ATTR_PRICE_OVERRIDE, false)
          sig = bounds_signature(entity)
          entity.set_attribute(DICT, ATTR_LAST_CALC_BOUNDS, sig.to_s) if sig
        end
        count_ref[0] += 1
      end
      apply_csv_to_tree(it[:children] || [], csv_data, count_ref) if it[:children]&.any?
    end
  end

  def self.merge_edits_into_csv_rows(rows, edits_json)
    return rows if edits_json.to_s.strip.empty?
    edits = JSON.parse(edits_json)
    rows.map do |r|
      id = (r[0] || "").to_s
      data = edits[id] || edits[id.to_i.to_s]
      next r unless data.is_a?(Hash)
      [
        r[0], r[1],
        (data["name"] || data[:name] || "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || r[2],
        r[3], r[4], r[5],
        (data["material_id"] || data[:material_id] || "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || (r[6] || ""),
        (data["thickness_mm"] || data[:thickness_mm] || "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || (r[7] || ""),
        (data["finish_id"] || data[:finish_id] || "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || (r[8] || ""),
        (data["material_name"] || data[:material_name] || "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || (r[9] || ""),
        (data["price"] || data[:price]).to_f.round,
        (data["memo"] || data[:memo] || "").to_s.strip.then { |s| (t = s.to_s.strip).empty? ? nil : t } || (r[11] || "")
      ]
    end
  rescue StandardError
    rows
  end

  def self.flatten_items_for_csv(items, _indent_level)
    result = []
    items.each do |it|
      next if it[:hidden]
      shape = (it[:children]&.any? ? "入れ子" : shape_type_label(it[:shape_type]))
      result << [
        (it[:id] || "").to_s, (it[:tag] || "").to_s, (it[:name] || "").to_s,
        (it[:type] || "").to_s, shape, (it[:detail] || "").to_s,
        (it[:material_id] || "").to_s, (it[:thickness_mm] || "").to_s, (it[:finish_id] || "").to_s, (it[:material_name] || "").to_s,
        (it[:price] || 0).to_f.round, (it[:memo] || "").to_s
      ]
      result.concat(flatten_items_for_csv(it[:children] || [], 0))
    end
    result
  end

  def self.shape_type_label(st)
    case st
    when "nested" then "入れ子"
    when "beam" then "棒"
    when "board" then "面状"
    when "freeform" then "自由形状"
    when "other" then "その他"
    else "—"
    end
  end

  # 形状変化した項目のID一覧（葉のみ。last_calc_bounds と現在 bounds を比較）
  def self.detect_shape_changed_ids
    result = []
    collect_ids_shape_changed = lambda do |items|
      items.each do |it|
        next if it[:hidden]
        if it[:children]&.any?
          collect_ids_shape_changed.call(it[:children])
          next
        end
        next if (it[:price] || 0).to_f <= 0
        entity = find_entity_by_pid(it[:id])
        next unless entity
        stored = entity.get_attribute(DICT, ATTR_LAST_CALC_BOUNDS, "").to_s.strip
        next if stored.empty?
        next unless stored.match?(/^\d+,\d+,\d+$/)
        current = bounds_signature(entity)
        next unless current
        result << it[:id].to_s if current != stored
      end
    end
    collect_ids_shape_changed.call(list_items)
    result
  end

  # ユーザーが手動で価格を上書きした項目のID一覧
  def self.detect_user_override_ids
    result = []
    collect_ids_user_override = lambda do |items|
      items.each do |it|
        next if it[:hidden]
        collect_ids_user_override.call(it[:children] || [])
        entity = find_entity_by_pid(it[:id])
        next unless entity
        result << it[:id].to_s if entity.get_attribute(DICT, ATTR_PRICE_OVERRIDE, false) == true
      end
    end
    collect_ids_user_override.call(list_items)
    result
  end

  def self.refresh_panel(highlight_ids: [], selected_ids: nil, expanded_ids: nil)
    return unless @dialog && @dialog.visible?
    propagate_nested_prices!(list_items)
    items = list_items.map { |it| serialize_item_for_panel(it) }
    raw_beam = raw_beam_options
    raw_board = raw_board_options
    finish_list = finish_by_id.values
    thickness_opts = board_thickness_options.map { |t| t.to_s }
    payload = {
      items: items, highlightIds: highlight_ids || [],
      shapeChangedIds: detect_shape_changed_ids,
      userOverrideIds: detect_user_override_ids,
      rawBeam: raw_beam, rawBoard: raw_board, finishList: finish_list,
      thicknessOptions: thickness_opts,
      version: VERSION
    }
    payload[:selectedIds] = selected_ids if selected_ids
    payload[:expandedIds] = expanded_ids if expanded_ids
    @dialog.execute_script("onListUpdated(#{payload.to_json});")
    show_panel_status("リスト: #{items.size}件")
  rescue => e
    show_panel_status("エラー: #{e.message}")
    UI.messagebox("パネル更新エラー\n\n#{format_error(e, 'refresh_panel')}")
  end

  def self.show_panel_status(msg)
    return unless @dialog && @dialog.visible?
    @dialog.execute_script("showStatus(#{msg.to_json});")
  rescue => e
  end

  # 複製（属性を転写）- ComponentInstance のみ対応（add_instance でコピー）
  def self.duplicate_with_attributes(ids_json)
    return unless model
    ids = ids_json.is_a?(String) ? JSON.parse(ids_json) : (ids_json || [])
    return if ids.empty?
    model.start_operation("複製（属性付き）", true)
    begin
      offset = Geom::Vector3d.new(100, 100, 0)
      copied = 0
      ids.each do |id_str|
        entity = find_entity_by_pid(id_str)
        next unless entity
        next unless entity.is_a?(Sketchup::ComponentInstance)
        item = find_item_in_tree(list_items, id_str.to_s)
        next unless item
        parent = entity.parent
        next unless parent.respond_to?(:entities)
        tr = entity.transformation * Geom::Transformation.new(offset)
        copy_ent = parent.entities.add_instance(entity.definition, tr)
        copy_ent.set_attribute(DICT, ATTR_MATERIAL_ID, (item[:material_id] || "").to_s)
        copy_ent.set_attribute(DICT, ATTR_THICKNESS_MM, (item[:thickness_mm] || "").to_s)
        copy_ent.set_attribute(DICT, ATTR_FINISH_ID, (item[:finish_id] || "").to_s)
        copy_ent.set_attribute(DICT, ATTR_MATERIAL_NAME, (item[:material_name] || "").to_s)
        copy_ent.set_attribute(DICT, ATTR_NAME, (item[:name] || "").to_s)
        copy_ent.set_attribute(DICT, ATTR_PRICE, ((item[:price] || 0).to_f).round)
        copy_ent.set_attribute(DICT, ATTR_MEMO, (item[:memo] || "").to_s)
        copied += 1
      end
      model.commit_operation
      refresh_list_from_model
      msg = copied > 0 ? "コンポーネント #{copied} 件を複製しました。" : "コンポーネントのみ複製可能です。グループはスキップしました。"
      UI.messagebox(msg)
    rescue => e
      model.abort_operation
      UI.messagebox("複製エラー\n#{e.message}")
    end
  end

  def self.reset_price(id)
    return unless model
    item = find_item_in_tree(list_items, id.to_s)
    return unless item
    entity = find_entity_by_pid(id)
    return unless entity
    return unless %w[beam board].include?(item[:shape_type])
    calc = compute_price_for_item(item, entity)
    return unless calc && calc > 0
    price_int = calc.round
    update_item_in_tree(list_items, id.to_s, :price, price_int)
    entity.set_attribute(DICT, ATTR_PRICE, price_int)
    entity.set_attribute(DICT, ATTR_PRICE_OVERRIDE, false)
    sig = bounds_signature(entity)
    entity.set_attribute(DICT, ATTR_LAST_CALC_BOUNDS, sig.to_s) if sig
    refresh_panel
  end

  def self.register_dialog_callbacks
    return unless @dialog
    @dialog.add_action_callback("clear_list") { clear_list; refresh_panel }
    @dialog.add_action_callback("toggle_visibility") { |_ctx, id| toggle_item_visibility(id) }
    @dialog.add_action_callback("zoom_to_item") { |_ctx, id| zoom_to_item(id) }
    @dialog.add_action_callback("refresh_list") { refresh_list_from_model }
    @dialog.add_action_callback("csv_load") { load_csv_from_file }
    @dialog.add_action_callback("update_price") do |_ctx, payload|
      next unless payload.is_a?(String) && payload.to_s.strip.start_with?("{")
      h = JSON.parse(payload)
      update_item_in_tree(list_items, h["id"], :price, ((h["price"] || 0).to_f).round)
      entity = find_entity_by_pid(h["id"])
      if entity&.respond_to?(:set_attribute)
        entity.set_attribute(DICT, ATTR_PRICE, ((h["price"] || 0).to_f).round)
        entity.set_attribute(DICT, ATTR_PRICE_OVERRIDE, true)
        sig = bounds_signature(entity)
        entity.set_attribute(DICT, ATTR_LAST_CALC_BOUNDS, sig.to_s) if sig
      end
      propagate_nested_prices!(list_items)
      refresh_panel
    end
    @dialog.add_action_callback("update_memo") do |_ctx, payload|
      next unless payload.is_a?(String) && payload.to_s.strip.start_with?("{")
      h = JSON.parse(payload)
      update_item_in_tree(list_items, h["id"], :memo, (h["memo"] || "").to_s)
      entity = find_entity_by_pid(h["id"])
      entity&.set_attribute(DICT, ATTR_MEMO, (h["memo"] || "").to_s)
    end
    @dialog.add_action_callback("update_name") do |_ctx, payload|
      next unless payload.is_a?(String) && payload.to_s.strip.start_with?("{")
      h = JSON.parse(payload)
      update_item_in_tree(list_items, h["id"], :name, (h["name"] || "").to_s.strip)
      entity = find_entity_by_pid(h["id"])
      entity&.set_attribute(DICT, ATTR_NAME, (h["name"] || "").to_s.strip)
    end
    @dialog.add_action_callback("update_material") do |_ctx, payload|
      next unless payload.is_a?(String) && payload.to_s.strip.start_with?("{")
      h = JSON.parse(payload)
      update_item_in_tree(list_items, h["id"], :material_id, (h["material_id"] || "").to_s)
      entity = find_entity_by_pid(h["id"])
      entity&.set_attribute(DICT, ATTR_MATERIAL_ID, (h["material_id"] || "").to_s)
    end
    @dialog.add_action_callback("update_thickness") do |_ctx, payload|
      next unless payload.is_a?(String) && payload.to_s.strip.start_with?("{")
      h = JSON.parse(payload)
      update_item_in_tree(list_items, h["id"], :thickness_mm, (h["thickness_mm"] || "").to_s)
      entity = find_entity_by_pid(h["id"])
      entity&.set_attribute(DICT, ATTR_THICKNESS_MM, (h["thickness_mm"] || "").to_s)
    end
    @dialog.add_action_callback("update_finish") do |_ctx, payload|
      next unless payload.is_a?(String) && payload.to_s.strip.start_with?("{")
      h = JSON.parse(payload)
      update_item_in_tree(list_items, h["id"], :finish_id, (h["finish_id"] || "").to_s)
      entity = find_entity_by_pid(h["id"])
      entity&.set_attribute(DICT, ATTR_FINISH_ID, (h["finish_id"] || "").to_s)
    end
    @dialog.add_action_callback("update_shape_type") do |_ctx, payload|
      next unless payload.is_a?(String) && payload.to_s.strip.start_with?("{")
      h = JSON.parse(payload)
      update_shape_type(h["id"], h["shape_type"])
    end
    @dialog.add_action_callback("update_material_name") do |_ctx, payload|
      next unless payload.is_a?(String) && payload.to_s.strip.start_with?("{")
      h = JSON.parse(payload)
      update_item_in_tree(list_items, h["id"], :material_name, (h["material_name"] || "").to_s)
      entity = find_entity_by_pid(h["id"])
      entity&.set_attribute(DICT, ATTR_MATERIAL_NAME, (h["material_name"] || "").to_s)
    end
    @dialog.add_action_callback("reset_price") { |_ctx, id| reset_price(id) }
    @dialog.add_action_callback("check_update") { check_for_updates(manual: true) }
    @dialog.add_action_callback("save_project") { |_ctx, edits_json| save_project(edits_json) }
    @dialog.add_action_callback("save_project_as") { |_ctx, edits_json| save_project_as(edits_json) }
    @dialog.add_action_callback("save_csv") { |_ctx, edits_json| save_csv(edits_json) }
    @dialog.add_action_callback("reload_materials_csv") { reload_materials_from_project_folder }
    @dialog.add_action_callback("reassign_from_shape") do |_ctx, include_overrides_str|
      inc = include_overrides_str.to_s.strip.downcase == "true"
      result = reassign_from_shape(include_overrides: inc)
      refresh_panel
      json = result.to_json
      @dialog.execute_script("if(typeof onRecalculateComplete==='function')onRecalculateComplete(#{json});")
    end
    @dialog.add_action_callback("get_selection_stats") do |_ctx, ids_json|
      ids = ids_json.is_a?(String) ? JSON.parse(ids_json) : (ids_json || [])
      propagate_nested_prices!(list_items)
      stats = compute_beam_board_stats(ids.map(&:to_s))
      @dialog.execute_script("onSelectionStatsUpdated(#{stats.to_json});")
    end
  end

  def self.show_panel
    html_path = panel_html_path
    unless File.exist?(html_path)
      UI.messagebox("パネルファイルが見つかりません:\n#{html_path}")
      return
    end
    @dialog ||= UI::HtmlDialog.new(
      dialog_title: EXT_NAME,
      preferences_key: "m_list_panel",
      width: 320,
      height: 560,
      resizable: true,
      style: UI::HtmlDialog::STYLE_UTILITY
    )
    @dialog.set_file(html_path)
    # HtmlDialog は再利用時にコールバックが失われるため、表示のたびに再登録
    register_dialog_callbacks
    @dialog.show
    @dialog.bring_to_front
    ensure_selection_observer
    refresh_list_from_model
    # バックグラウンドでアップデート確認（2週間ごと）
    UI.start_timer(2, false) { check_for_updates(manual: false) }
  end

  class SelectionSyncObserver < Sketchup::SelectionObserver
    def onSelectionBulkChange(selection)
      MList.sync_list_from_selection
    end

    def onSelectionCleared(selection)
      MList.sync_list_from_selection
    end
  end

  def self.ensure_selection_observer
    return unless model && model.selection
    @_selection_observer_added ||= {}
    return if @_selection_observer_added[model.object_id]
    model.selection.add_observer(SelectionSyncObserver.new)
    @_selection_observer_added[model.object_id] = true
  end

  class AppObserver < Sketchup::AppObserver
    def expectsStartupModelNotifications
      true
    end

    def onActivateModel(active_model)
      handle_model_change(active_model)
    end

    def onOpenModel(opened_model)
      handle_model_change(opened_model)
    end

    def onNewModel(new_model)
      handle_model_change(new_model)
    end

    def handle_model_change(active_model)
      return unless active_model
      MList.instance_variable_set(:@list_items, [])
      MList.instance_variable_set(:@csv_path, nil)
      MList.instance_variable_set(:@raw_by_id, nil)
      MList.instance_variable_set(:@finish_by_id, nil)
      MList.refresh_list_from_model
      MList.ensure_selection_observer
      MList.refresh_panel if MList.instance_variable_get(:@dialog)&.visible?
    end
  end

  unless file_loaded?(__FILE__)
    Sketchup.add_observer(AppObserver.new)
    ext_menu = UI.menu("Extensions")
    ext_menu.add_item(EXT_NAME) { show_panel }
    ext_menu.add_item("#{EXT_NAME} - アップデートを確認") { check_for_updates(manual: true) }
    file_loaded(__FILE__)
  end
end
