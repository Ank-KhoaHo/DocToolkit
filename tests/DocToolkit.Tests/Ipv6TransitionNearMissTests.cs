using System.Net;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// Near misses on the IPv6 transition prefixes: addresses that match <b>part</b> of a prefix and
/// must NOT be unwrapped as one.
///
/// <c>GuardedResourceLoaderTests</c> proves the positive direction — a genuine 6to4, Teredo or
/// NAT64 wrapper around a private address is blocked. That is one half of the property. The other
/// half is that the byte tests are <i>conjunctions</i>: every byte must match. Mutation testing
/// found that half untested, surviving <c>&amp;&amp;</c> → <c>||</c> on all five prefix checks in
/// <c>TryGetEmbeddedIPv4</c>.
///
/// Each address below is a global-unicast address that is legitimately reachable today. Under a
/// loosened conjunction it would instead be unwrapped, reveal 10.0.0.5, and be refused — so a
/// caller would silently lose access to real hosts. The failure a loosened guard causes is
/// over-blocking, not under-blocking, which is why nothing else notices it.
/// </summary>
public class Ipv6TransitionNearMissTests
{
    [Theory]
    // NAT64 well-known, 64:ff9b::/96 — requires bytes 4-11 to be zero. Here they are not, so this is
    // an ordinary address that happens to share the 64:ff9b prefix.
    [InlineData("64:ff9b:203:405:607:809:a00:5")]
    // NAT64 local-use, 64:ff9b:1::/48 — requires b[4]=0x00 AND b[5]=0x01. Here b[5] is 0x02.
    [InlineData("64:ff9b:2:a00:0:500::")]
    // 6to4, 2002::/16 — requires b[0]=0x20 AND b[1]=0x02. Here b[1] is 0x03.
    [InlineData("2003:a00:5::")]
    // Teredo, 2001::/32 — requires b[0..3] = 20 01 00 00. Here b[3] is 0x01.
    // Bytes 12-15 are 10.0.0.5 XOR 0xFF, so a loosened check decodes them straight to 10.0.0.5.
    [InlineData("2001:1::f5ff:fffa")]
    public void AnAddressMatchingOnlyPartOfATransitionPrefixIsNotUnwrapped(string address)
    {
        Assert.False(
            GuardedResourceLoader.IsBlockedAddress(IPAddress.Parse(address)),
            $"{address} matches only part of a transition prefix and must be treated as the "
            + "ordinary global-unicast address it is. Blocking it means a byte test was loosened "
            + "from AND to OR, and real hosts are now unreachable.");
    }

    // The positive counterpart, kept beside the near misses so the pair reads as one statement:
    // change one byte and the same address IS a wrapper around 10.0.0.5, and IS refused. Without
    // this, every assertion above would still pass against a guard that unwrapped nothing at all.
    [Theory]
    [InlineData("64:ff9b::a00:5")]          // NAT64 well-known, bytes 4-11 zero
    [InlineData("64:ff9b:1:a00:0:500::")]   // NAT64 local-use, b[5] = 0x01
    [InlineData("2002:a00:5::")]            // 6to4, b[1] = 0x02
    [InlineData("2001::f5ff:fffa")]         // Teredo, b[3] = 0x00
    public void TheSameAddressWithTheFullPrefixIsBlocked(string address)
    {
        Assert.True(
            GuardedResourceLoader.IsBlockedAddress(IPAddress.Parse(address)),
            $"{address} is a transition wrapper around 10.0.0.5 and must be refused. If this "
            + "passes while the near-miss tests also pass, the guard is unwrapping nothing.");
    }
}
