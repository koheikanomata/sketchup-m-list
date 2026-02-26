# frozen_string_literal: true
# M List - 価格計算・規格品必要数

module EstimateAuto
  def self.compute_beam_board_stats(selected_ids)
    result = { beam_count: 0, beam_total: 0.0, board_count: 0, board_area: 0.0, total_price: 0.0, breakdown: [] }
    seen = {}
    beam_by_material = {}
    board_by_material = {}
    finish_by_usage = {}
    material_seen = {}
    selected_ids.uniq.each do |id|
      item = find_item_in_tree(list_items, id.to_s)
      next unless item
      next if item[:hidden]
      r = collect_beam_board_from_item(item, seen)
      result[:beam_count] += r[:beam_count]
      result[:beam_total] += r[:beam_total]
      result[:board_count] += r[:board_count]
      result[:board_area] += r[:board_area]
      result[:total_price] += (item[:price] || 0).to_f
      result[:breakdown] << build_calc_detail(item)
      collect_material_usage(item, material_seen, beam_by_material, board_by_material, finish_by_usage)
    end
    board_specs = compute_board_specs(board_by_material)
    beam_specs = compute_beam_specs(beam_by_material)
    finish_specs = compute_finish_specs(finish_by_usage)
    {
      beam_count: result[:beam_count],
      beam_total_mm: result[:beam_total].round(1),
      board_count: result[:board_count],
      board_area_mm2: result[:board_area].round(1),
      total_price: result[:total_price].round,
      breakdown: result[:breakdown],
      board_specs: board_specs,
      beam_specs: beam_specs,
      finish_specs: finish_specs
    }
  end

  def self.collect_material_usage(item, seen, beam_by_material, board_by_material, finish_by_usage)
    return if item.nil?
    return if seen[item[:id].to_s]
    return if item[:hidden]
    seen[item[:id].to_s] = true
    if item[:children]&.any?
      item[:children].each do |c|
        collect_material_usage(c, seen, beam_by_material, board_by_material, finish_by_usage) unless c.nil? || c[:hidden]
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
      { material_id: mid, total_m2: total_m2.round(4), sheet_m2: sheet_m2.round(4), sheets_needed: sheets_needed }
    end
  end

  def self.compute_beam_specs(beam_by_material)
    beam_by_material.map do |mid, h|
      total_mm = (h[:total_mm] || 0).round(1)
      unit_len = h[:unit_len_mm] || 1000
      beams_needed = unit_len > 0 ? (total_mm / unit_len).ceil : 0
      { material_id: mid, total_mm: total_mm, unit_len_mm: unit_len.round, beams_needed: beams_needed }
    end
  end

  def self.compute_finish_specs(finish_by_usage)
    finish_by_usage.map do |fid, total_mm2|
      total_m2 = total_mm2 / 1_000_000.0
      fin = finish_by_id[fid]
      coverage = fin && fin["m2_coverage_per_unit"].to_f > 0 ? fin["m2_coverage_per_unit"].to_f : 1.0
      units_needed = (total_m2 / coverage).ceil
      { finish_id: fid, total_m2: total_m2.round(4), coverage_per_unit: coverage, units_needed: units_needed }
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
      area_mm2 = entity ? board_area_mm2(entity) : 0
      area_m2 = area_mm2 / 1_000_000.0
      raw = entity ? raw_by_id[item[:material_id].to_s.strip] : nil
      fin = finish_by_id[item[:finish_id].to_s.strip]
      price_per_m2 = 0.0
      if raw && sheet_unit?(raw["unit_size"])
        m = raw["unit_size"].to_s.match(/(\d+)\s*[xX×]\s*(\d+)/)
        if m
          sheet_area = (m[1].to_f / 1000.0) * (m[2].to_f / 1000.0)
          price_per_m2 += (sheet_area > 0 ? (raw["unit_price"].to_f / sheet_area) * (1 + raw["loss_rate"].to_f / 100.0) : 0)
        end
      end
      price_per_m2 += fin["1m2_price"].to_f if fin
      if price_per_m2 > 0
        { name: name, price: price, detail: "面積 #{area_m2.round(4)} m² × ¥#{price_per_m2.round}/m² = ¥#{price}" }
      else
        { name: name, price: price, detail: "直接入力" }
      end
    else
      { name: name, price: price, detail: "直接入力" }
    end
  end

  def self.collect_beam_board_from_item(item, seen)
    return { beam_count: 0, beam_total: 0.0, board_count: 0, board_area: 0.0 } if item.nil?
    return { beam_count: 0, beam_total: 0.0, board_count: 0, board_area: 0.0 } if seen[item[:id].to_s]
    return { beam_count: 0, beam_total: 0.0, board_count: 0, board_area: 0.0 } if item[:hidden]
    seen[item[:id].to_s] = true
    if item[:children]&.any?
      beam_c, beam_t, board_c, board_a = 0, 0.0, 0, 0.0
      item[:children].each do |c|
        next if c.nil? || c[:hidden]
        r = collect_beam_board_from_item(c, seen)
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
      area_m2 = board_area_mm2(entity) / 1_000_000.0
      total = 0.0

      # 下地材の価格（選択時のみ）
      raw = raw_by_id[item[:material_id].to_s.strip]
      if raw && sheet_unit?(raw["unit_size"])
        m = raw["unit_size"].to_s.match(/(\d+)\s*[xX×]\s*(\d+)/)
        if m
          sheet_area = (m[1].to_f / 1000.0) * (m[2].to_f / 1000.0)
          if sheet_area > 0
            price_per_m2 = (raw["unit_price"].to_f / sheet_area) * (1 + raw["loss_rate"].to_f / 100.0)
            total += price_per_m2 * area_m2
          end
        end
      end

      # 仕上材の価格（選択時のみ）
      fin = finish_by_id[item[:finish_id].to_s.strip]
      total += (fin["1m2_price"].to_f * area_m2) if fin

      total > 0 ? total : nil
    else
      nil
    end
  end
end
