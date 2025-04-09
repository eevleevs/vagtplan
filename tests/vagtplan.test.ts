import { assertEquals } from '@std/assert'

import { extension, main } from '../vagtplan.ts'

for (const { name } of Deno.readDirSync('tests')) {
  if (name.startsWith('~$') || !name.match(extension)) continue
  Deno.test(name, () =>
    assertEquals(
      main(`tests/${name}`),
      JSON.parse(
        Deno.readTextFileSync(`tests/${name}`.replace(extension, '.json')),
      ),
    ))
}
