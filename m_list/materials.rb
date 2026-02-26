# frozen_string_literal: true
# M List - 材料マスタ読み込み・検索

module EstimateAuto
  def self.path_materials_board
    File.join(plugin_root, "materials_board.csv")
  end

  def self.path_materials_beam
    File.join(plugin_root, "materials_beam.csv")
  end

  def self.path_materials_finish
    File.join(plugin_root, "materials_finish.csv")
  end

  def self.load_materials_board
    return [] unless File.exist?(path_materials_board)
    CSV.read(path_materials_board, headers: true, encoding: "UTF-8").map(&:to_h)
  rescue StandardError
    []
  end

  def self.load_materials_beam
    return [] unless File.exist?(path_materials_beam)
    CSV.read(path_materials_beam, headers: true, encoding: "UTF-8").map(&:to_h)
  rescue StandardError
    []
  end

  def self.load_materials_finish
    return [] unless File.exist?(path_materials_finish)
    CSV.read(path_materials_finish, headers: true, encoding: "UTF-8").map(&:to_h)
  rescue StandardError
    []
  end

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

  # 板材CSV行から厚み(mm)を抽出。name の "12mm" や "2.5mm" をパース
  def self.parse_board_thickness_mm(row)
    return nil unless row
    name = (row["name"] || "").to_s
    m = name.match(/(\d+\.?\d*)\s*mm/i)
    m ? m[1].to_f : nil
  end

  # 板材エンティティの厚み(mm)＝boundsの最小寸法
  def self.board_thickness_mm(entity)
    return nil unless entity&.respond_to?(:bounds)
    bb = entity.bounds
    dims = [bb.width.to_mm, bb.height.to_mm, bb.depth.to_mm].sort
    dims[0] > 0 ? dims[0] : nil
  end

  # 厚みが最も近い板材の material_id を返す（同一厚み優先）
  def self.find_best_board_match_by_thickness(entity)
    target_mm = board_thickness_mm(entity)
    return nil unless target_mm && target_mm > 0
    candidates = raw_board_options.map do |r|
      row_mm = parse_board_thickness_mm(r)
      next nil unless row_mm && row_mm > 0
      diff = (row_mm - target_mm).abs
      { material_id: r["material_id"].to_s.strip, diff: diff, exact: (diff < 0.01) }
    end.compact
    return nil if candidates.empty?
    exact = candidates.select { |c| c[:exact] }
    best = exact.any? ? exact.min_by { |c| c[:diff] } : candidates.min_by { |c| c[:diff] }
    best[:material_id]
  end

  def self.sheet_unit?(size_str)
    return false unless size_str
    size_str.to_s.match?(/^\d+\s*[xX×]\s*\d+$/)
  end

  def self.linear_unit?(size_str)
    return false unless size_str
    size_str.to_s.match?(/^\d+\s*[xX×]\s*\d+\s*[xX×]\s*\d+$/)
  end
end
