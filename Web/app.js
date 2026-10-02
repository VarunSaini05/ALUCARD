const repository = 'https://github.com/VarunSaini05/ALUCARD';
const latestReleaseApi = 'https://api.github.com/repos/VarunSaini05/ALUCARD/releases/latest';
const installerName = 'ALUCARD-Setup.exe';

const downloadLinks = document.querySelectorAll('.js-download');
const downloadStates = document.querySelectorAll('[data-download-state]');
const releaseState = document.querySelector('[data-release-state]');
const versionValue = document.querySelector('[data-version]');
const sizeValue = document.querySelector('[data-size]');
const dateValue = document.querySelector('[data-date]');
const releaseLink = document.querySelector('[data-release-link]');

function setDownloadUnavailable(message) {
  for (const link of downloadLinks) {
    link.removeAttribute('href');
    link.setAttribute('aria-disabled', 'true');
    link.setAttribute('tabindex', '-1');
  }

  for (const state of downloadStates) {
    state.textContent = message;
  }
}

function formatFileSize(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return 'Not listed';
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function formatReleaseDate(dateString) {
  if (!dateString) return 'Not listed';
  const date = new Date(dateString);
  if (Number.isNaN(date.getTime())) return 'Not listed';

  return new Intl.DateTimeFormat('en', {
    day: 'numeric',
    month: 'short',
    year: 'numeric',
    timeZone: 'UTC'
  }).format(date);
}

async function loadLatestRelease() {
  let unavailableMessage = 'Release information could not be loaded. Please try again later.';

  try {
    const response = await fetch(latestReleaseApi, {
      headers: { Accept: 'application/vnd.github+json' },
      cache: 'no-store'
    });

    if (response.status === 404) {
      unavailableMessage = 'Latest installer is being prepared.';
    }

    if (!response.ok) {
      throw new Error(`Release API returned ${response.status}`);
    }

    const release = await response.json();
    const installer = release.assets?.find((asset) => asset.name === installerName);

    if (!installer) {
      unavailableMessage = 'Latest installer is being prepared.';
      throw new Error('The latest release does not contain the installer asset');
    }

    versionValue.textContent = release.tag_name || release.name || 'Latest';
    sizeValue.textContent = formatFileSize(installer.size);
    dateValue.textContent = formatReleaseDate(release.published_at);
    releaseLink.href = release.html_url || `${repository}/releases`;

    const downloadUrl = `${repository}/releases/latest/download/${installerName}`;
    for (const link of downloadLinks) {
      link.href = downloadUrl;
      link.removeAttribute('aria-disabled');
      link.removeAttribute('tabindex');
    }

    for (const state of downloadStates) {
      state.textContent = `${versionValue.textContent} is available`;
    }
    releaseState.textContent = 'Latest installer verified from GitHub Releases.';
  } catch (error) {
    versionValue.textContent = 'Not available';
    sizeValue.textContent = 'Not available';
    dateValue.textContent = 'Not available';
    releaseState.textContent = unavailableMessage;
    setDownloadUnavailable(unavailableMessage);
  }
}

loadLatestRelease();