import { cp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import path from 'node:path'

const root = process.cwd()
const bundleDir = path.join(root, 'bundle')

async function main() {
  await rm(bundleDir, { recursive: true, force: true })
  await mkdir(bundleDir, { recursive: true })

  await cp(path.join(root, 'dist'), path.join(bundleDir, 'dist'), {
    recursive: true,
  })
  await cp(path.join(root, 'dist-electron'), path.join(bundleDir, 'dist-electron'), {
    recursive: true,
  })

  const packageJson = JSON.parse(
    await readFile(path.join(root, 'package.json'), 'utf8'),
  )

  const minimalPackageJson = {
    name: packageJson.name,
    version: packageJson.version,
    description: packageJson.description,
    author: packageJson.author,
    main: packageJson.main,
  }

  await writeFile(
    path.join(bundleDir, 'package.json'),
    `${JSON.stringify(minimalPackageJson, null, 2)}\n`,
    'utf8',
  )
}

await main()
