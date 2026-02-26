# frozen_string_literal: true
# M List - 形状判定

module EstimateAuto
  def self.collect_children(entity)
    entities = entity.is_a?(Sketchup::Group) ? entity.entities : entity.definition.entities
    entities.select { |e| e.is_a?(Sketchup::Group) || e.is_a?(Sketchup::ComponentInstance) }
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

  def self.board_area_mm2(entity)
    return 0 unless entity.respond_to?(:bounds)
    dims = [entity.bounds.width.to_mm, entity.bounds.height.to_mm, entity.bounds.depth.to_mm].sort
    return 0 if dims[1] <= 0 || dims[2] <= 0
    dims[1] * dims[2]
  end

  def self.determine_shape_type(e)
    return "nested" if collect_children(e).any?
    return "beam" if beam_like?(e)
    return "board" if board_like?(e)
    return "freeform" if face_count(e) >= FREEFORM_MIN_FACE_COUNT
    "other"
  end
end
