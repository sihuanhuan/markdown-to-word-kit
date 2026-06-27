-- Pandoc Lua filter for Mermaid and PlantUML fenced code blocks.
-- Requires `mmdc` for Mermaid and `plantuml` or `PLANTUML_JAR` for PlantUML.

local diagram_dir = ".pandoc-diagrams"
local is_windows = package.config:sub(1, 1) == "\\"

local function shell_quote(value)
  value = tostring(value)
  if is_windows then
    return '"' .. value:gsub('"', '\\"') .. '"'
  end
  return "'" .. value:gsub("'", "'\\''") .. "'"
end

local function command_exists(command)
  local probe
  if is_windows then
    probe = "where " .. shell_quote(command) .. " >nul 2>nul"
  else
    probe = "command -v " .. shell_quote(command) .. " >/dev/null 2>&1"
  end
  local ok = os.execute(probe)
  return ok == true or ok == 0
end

local function write_file(path, text)
  local file = assert(io.open(path, "w"))
  file:write(text)
  file:close()
end

local function run(command, message)
  local ok = os.execute(command)
  if not (ok == true or ok == 0) then
    error(message .. "\nCommand failed: " .. command)
  end
end

local function ensure_dir(path)
  if is_windows then
    run("if not exist " .. shell_quote(path) .. " mkdir " .. shell_quote(path), "Unable to create diagram output directory.")
  else
    run("mkdir -p " .. shell_quote(path), "Unable to create diagram output directory.")
  end
end

local function has_class(el, name)
  for _, class in ipairs(el.classes) do
    if class == name then
      return true
    end
  end
  return false
end

local function image_for(el, kind, target)
  local caption = el.attributes.caption or kind .. " diagram"
  return pandoc.Para({ pandoc.Image(caption, target) })
end

function CodeBlock(el)
  local is_mermaid = has_class(el, "mermaid")
  local is_plantuml = has_class(el, "plantuml") or has_class(el, "puml")

  if not is_mermaid and not is_plantuml then
    return nil
  end

  ensure_dir(diagram_dir)

  local hash = pandoc.utils.sha1(el.text)

  if is_mermaid then
    if not command_exists("mmdc") then
      error("Mermaid rendering needs `mmdc`. Run install-tools.ps1 -WithMermaid or install @mermaid-js/mermaid-cli and add mmdc to PATH.")
    end

    local input = diagram_dir .. "/" .. hash .. ".mmd"
    local output = diagram_dir .. "/" .. hash .. ".png"
    write_file(input, el.text)
    run(
      "mmdc -i " .. shell_quote(input) .. " -o " .. shell_quote(output) .. " -b transparent",
      "Mermaid rendering failed."
    )
    return image_for(el, "Mermaid", output)
  end

  if is_plantuml then
    local input = diagram_dir .. "/" .. hash .. ".puml"
    local output = diagram_dir .. "/" .. hash .. ".png"
    write_file(input, el.text)

    if command_exists("plantuml") then
      run(
        "plantuml -tpng -pipe < " .. shell_quote(input) .. " > " .. shell_quote(output),
        "PlantUML rendering failed."
      )
    else
      local jar = os.getenv("PLANTUML_JAR")
      if jar == nil or jar == "" then
        error("PlantUML rendering needs `plantuml` or `PLANTUML_JAR=/path/to/plantuml.jar`.")
      end
      run(
        "java -jar " .. shell_quote(jar) .. " -tpng -pipe < " .. shell_quote(input) .. " > " .. shell_quote(output),
        "PlantUML rendering failed."
      )
    end

    return image_for(el, "PlantUML", output)
  end
end
